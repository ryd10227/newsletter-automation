[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)][string]$Template,
  [Parameter(Mandatory = $true)][string]$Csv,
  [Parameter(Mandatory = $true)][string]$Psd,
  [Parameter(Mandatory = $true)][string]$MapHtml,
  [Parameter(Mandatory = $true)][string]$Output,
  [string]$IssueDate,
  [switch]$BestEffort
)

$ErrorActionPreference = "Stop"
$Warnings = New-Object System.Collections.Generic.List[string]
$Capacities = @{ business = 8; academy = 4; banner = 3 }
$MarketLinkSlots = @(1, 3, 5, 2, 4, 6)
$AreaPattern = "<area\b[^>]*>"
$AnchorPattern = "<a\b[^>]*>.*?</a\s*>"
$ImagePattern = "<img\b[^>]*>"

function Get-CleanText {
  param($Value)
  if ($null -eq $Value) { return "" }
  return ([regex]::Replace(([string]$Value).Replace([char]0x00A0, " ").Replace("`r", " ").Replace("`n", " "), "\s+", " ")).Trim()
}

function ConvertTo-HtmlText {
  param([string]$Value)
  return [Net.WebUtility]::HtmlEncode($Value)
}

function Get-AttrValue {
  param([string]$Tag, [string]$Name)
  $Pattern = "\b$([regex]::Escape($Name))\s*=\s*([`"'])(.*?)\1"
  $Match = [regex]::Match($Tag, $Pattern, "IgnoreCase,Singleline")
  if (-not $Match.Success) { return "" }
  return [Net.WebUtility]::HtmlDecode($Match.Groups[2].Value).Trim()
}

function Set-AttrValue {
  param([string]$Tag, [string]$Name, [string]$Value)
  $Pattern = "\b$([regex]::Escape($Name))\s*=\s*([`"'])(.*?)\1"
  $Replacement = $Name + '="' + (ConvertTo-HtmlText $Value) + '"'
  $Regex = New-Object Text.RegularExpressions.Regex($Pattern, [Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [Text.RegularExpressions.RegexOptions]::Singleline)
  if ($Regex.IsMatch($Tag)) {
    return $Regex.Replace($Tag, [Text.RegularExpressions.MatchEvaluator]{ param($m) $Replacement }, 1)
  }
  return $Tag.Substring(0, $Tag.Length - 1).TrimEnd() + " $Replacement>"
}

function Read-CsvRecords {
  param([string]$Path)
  $Rows = @(Import-Csv -LiteralPath $Path -Encoding UTF8)
  if ($Rows.Count -eq 0) { throw "CSV has no data rows." }
  $Names = @($Rows[0].PSObject.Properties.Name)
  foreach ($Required in @("section", "slot", "title", "period", "deadline", "url")) {
    if ($Names -notcontains $Required) { throw "CSV requires section, slot, title, period, deadline, and url columns." }
  }
  $Records = @{}
  $CsvDate = $null
  $LineNo = 1
  foreach ($Row in $Rows) {
    $LineNo++
    $DateText = Get-CleanText $Row.newsletter_date
    if ($DateText -and $null -eq $CsvDate) {
      [DateTime]$ParsedDate = [DateTime]::MinValue
      if ([DateTime]::TryParseExact($DateText, "yyyy-MM-dd", [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$ParsedDate)) {
        $CsvDate = $ParsedDate.Date
      } else {
        $Message = "CSV line $LineNo has an invalid newsletter_date: $DateText"
        if ($BestEffort) { $Warnings.Add($Message) } else { throw $Message }
      }
    }
    $Section = (Get-CleanText $Row.section).ToLowerInvariant()
    if (-not $Capacities.ContainsKey($Section)) { continue }
    [int]$Slot = 0
    if (-not [int]::TryParse((Get-CleanText $Row.slot), [ref]$Slot)) {
      $Message = "CSV line $LineNo has a non-numeric slot."
      if ($BestEffort) { $Warnings.Add($Message); continue } else { throw $Message }
    }
    if ($Slot -lt 1 -or $Slot -gt $Capacities[$Section]) {
      $Message = "CSV line $LineNo exceeds the $Section capacity."
      if ($BestEffort) { $Warnings.Add($Message); continue } else { throw $Message }
    }
    $Key = "$Section|$Slot"
    if ($Records.ContainsKey($Key)) {
      $Message = "CSV contains duplicate $Section slot $Slot."
      if ($BestEffort) { $Warnings.Add($Message); continue } else { throw $Message }
    }
    $Record = @{}
    foreach ($Name in $Names) { $Record[$Name] = Get-CleanText $Row.$Name }
    if ($Record.url -and $Record.url -notmatch '^https?://') {
      $Message = "CSV line $LineNo has an invalid URL; it was cleared: $($Record.url)"
      if ($BestEffort) { $Warnings.Add($Message); $Record.url = "" } else { throw $Message }
    }
    $Records[$Key] = $Record
  }
  return [pscustomobject]@{ Records = $Records; Date = $CsvDate }
}

function Read-MapLinks {
  param([string]$Path)
  $Source = [IO.File]::ReadAllText($Path, [Text.Encoding]::UTF8)
  $Links = @{}
  foreach ($Match in [regex]::Matches($Source, $AreaPattern, "IgnoreCase,Singleline")) {
    $Alt = Get-AttrValue $Match.Value "alt"
    $Href = Get-AttrValue $Match.Value "href"
    if ($Alt -and $Href) { $Links[$Alt] = $Href }
  }
  return $Links
}

function Get-PrimarySectionOrder {
  param([string]$Path)
  $Source = [IO.File]::ReadAllText($Path, [Text.Encoding]::UTF8)
  $Market = $null
  $Business = $null
  foreach ($Match in [regex]::Matches($Source, $AreaPattern, "IgnoreCase,Singleline")) {
    $Alt = Get-AttrValue $Match.Value "alt"
    $Coordinates = @([regex]::Matches((Get-AttrValue $Match.Value "coords"), "-?\d+") | ForEach-Object { [int]$_.Value })
    $Top = if ($Coordinates.Count -ge 2) { $Coordinates[1] } else { $Match.Index }
    $Key = @($Top, $Match.Index)
    if ($null -eq $Market -and @("해외시장뉴스_메인", "헤드라인", "리스트1") -contains $Alt) { $Market = $Key }
    if ($null -eq $Business -and @("사업안내_메인", "사업안내1") -contains $Alt) { $Business = $Key }
    if ($null -ne $Market -and $null -ne $Business) { break }
  }
  if ($null -eq $Market -or $null -eq $Business) { throw "Could not determine the market/business order from the image map." }
  if ($Business[0] -lt $Market[0] -or ($Business[0] -eq $Market[0] -and $Business[1] -lt $Market[1])) { return "business_first" }
  return "market_first"
}

function Get-U16BE { param([byte[]]$Data, [int]$Offset) return [int]$Data[$Offset] * 256 + [int]$Data[$Offset + 1] }
function Get-I16BE {
  param([byte[]]$Data, [int]$Offset)
  $Value = Get-U16BE $Data $Offset
  if ($Value -ge 32768) { return $Value - 65536 }
  return $Value
}
function Get-U32BE {
  param([byte[]]$Data, [int]$Offset)
  return [uint32]([uint64]$Data[$Offset] * 16777216 + [uint64]$Data[$Offset + 1] * 65536 + [uint64]$Data[$Offset + 2] * 256 + [uint64]$Data[$Offset + 3])
}
function Get-I32BE {
  param([byte[]]$Data, [int]$Offset)
  [uint64]$Value = Get-U32BE $Data $Offset
  if ($Value -ge 2147483648) { return [int64]$Value - 4294967296 }
  return [int64]$Value
}
function Get-AsciiKey {
  param([byte[]]$Data, [int]$Offset, [int]$Length)
  return [Text.Encoding]::ASCII.GetString($Data, $Offset, $Length)
}

function Read-PsdLayers {
  param([string]$Path)
  [byte[]]$Data = [IO.File]::ReadAllBytes($Path)
  if ($Data.Length -lt 30 -or (Get-AsciiKey $Data 0 4) -ne "8BPS") { throw "Not a valid PSD file." }
  if ((Get-U16BE $Data 4) -ne 1) { throw "PSB is not supported; save as PSD." }
  [int64]$Position = 26
  for ($Index = 0; $Index -lt 2; $Index++) {
    [int64]$Length = Get-U32BE $Data ([int]$Position)
    $Position += 4 + $Length
  }
  [int64]$LayerMaskLength = Get-U32BE $Data ([int]$Position)
  $Position += 4
  if ($LayerMaskLength -eq 0) { throw "PSD has no layer information." }
  [int64]$LayerInfoLength = Get-U32BE $Data ([int]$Position)
  $Position += 4
  if ($LayerInfoLength -lt 2) { throw "PSD has no layer records." }
  $LayerCount = [Math]::Abs((Get-I16BE $Data ([int]$Position)))
  $Position += 2
  $Layers = @()
  for ($LayerIndex = 0; $LayerIndex -lt $LayerCount; $LayerIndex++) {
    $Top = Get-I32BE $Data ([int]$Position); $Position += 4
    $Left = Get-I32BE $Data ([int]$Position); $Position += 4
    $Bottom = Get-I32BE $Data ([int]$Position); $Position += 4
    $Right = Get-I32BE $Data ([int]$Position); $Position += 4
    $ChannelCount = Get-U16BE $Data ([int]$Position)
    $Position += 2 + $ChannelCount * 6
    $Position += 12
    [int64]$ExtraLength = Get-U32BE $Data ([int]$Position); $Position += 4
    [int64]$ExtraEnd = $Position + $ExtraLength
    [int64]$MaskLength = Get-U32BE $Data ([int]$Position); $Position += 4 + $MaskLength
    [int64]$RangesLength = Get-U32BE $Data ([int]$Position); $Position += 4 + $RangesLength
    $PascalLength = [int]$Data[[int]$Position]; $Position++
    $PascalName = ""
    if ($PascalLength -gt 0) {
      try { $PascalName = [Text.Encoding]::GetEncoding(10000).GetString($Data, [int]$Position, $PascalLength) }
      catch { $PascalName = [Text.Encoding]::Default.GetString($Data, [int]$Position, $PascalLength) }
    }
    $Position += $PascalLength
    $Position += (4 - ((1 + $PascalLength) % 4)) % 4
    $UnicodeName = $null
    $IsText = $false
    while ($Position + 12 -le $ExtraEnd) {
      $Signature = Get-AsciiKey $Data ([int]$Position) 4
      $Key = Get-AsciiKey $Data ([int]$Position + 4) 4
      [int64]$BlockLength = Get-U32BE $Data ([int]$Position + 8)
      $Position += 12
      if (@("8BIM", "8B64") -notcontains $Signature -or $Position + $BlockLength -gt $ExtraEnd) { break }
      if ($Key -eq "luni" -and $BlockLength -ge 4) {
        $CharCount = Get-U32BE $Data ([int]$Position)
        $ByteCount = [Math]::Min([int64]$CharCount * 2, $BlockLength - 4)
        $UnicodeName = [Text.Encoding]::BigEndianUnicode.GetString($Data, [int]$Position + 4, [int]$ByteCount)
      } elseif ($Key -eq "TySh") { $IsText = $true }
      $Position += $BlockLength
      if ($Position % 2 -ne 0) { $Position++ }
    }
    $Position = $ExtraEnd
    $Name = if ($UnicodeName) { $UnicodeName } else { $PascalName }
    $Layers += [pscustomobject]@{
      Index = $LayerIndex; Top = [int]$Top; Left = [int]$Left; Bottom = [int]$Bottom; Right = [int]$Right
      Name = Get-CleanText $Name; IsText = $IsText
    }
  }
  return $Layers
}

function Split-MarketLocation {
  param([string]$Value)
  $LocationRegex = New-Object Text.RegularExpressions.Regex("[ㅣ|]")
  $Parts = $LocationRegex.Split((Get-CleanText $Value), 2)
  if ($Parts.Count -ne 2) { return $null }
  $Country = Get-CleanText $Parts[0]
  $Office = Get-CleanText $Parts[1]
  if (-not $Country -or -not $Office -or $Office -notmatch '무역(?:관)?$') { return $null }
  if ($Office.EndsWith("무역")) { $Office += "관" }
  return [pscustomobject]@{ Country = $Country; Office = $Office }
}

function Assign-MarketSlots {
  param([object[]]$Items)
  if ($Items.Count -eq 0) { return @() }
  $RowTops = New-Object System.Collections.Generic.List[int]
  foreach ($Top in @($Items | ForEach-Object Top | Sort-Object)) {
    if ($RowTops.Count -eq 0 -or $Top - $RowTops[$RowTops.Count - 1] -gt 20) { $RowTops.Add($Top) }
  }
  if ($RowTops.Count -gt 3) { throw "Found $($RowTops.Count) market-news rows instead of 3." }
  $LeftEdge = ($Items | Measure-Object Left -Minimum).Minimum
  $RightEdge = ($Items | Measure-Object Left -Maximum).Maximum
  $Split = if ($RightEdge - $LeftEdge -gt 180) { ($LeftEdge + $RightEdge) / 2 } else { 300 }
  foreach ($Item in $Items) {
    $Nearest = 0
    for ($Index = 1; $Index -lt $RowTops.Count; $Index++) {
      if ([Math]::Abs($Item.Top - $RowTops[$Index]) -lt [Math]::Abs($Item.Top - $RowTops[$Nearest])) { $Nearest = $Index }
    }
    $Column = if ($Item.Left -lt $Split) { 0 } else { 1 }
    $Item.MapSlot = $Nearest * 2 + $Column + 1
  }
  return @($Items | Sort-Object { [Array]::IndexOf($MarketLinkSlots, [int]$_.MapSlot) })
}

function Extract-PsdContent {
  param([object[]]$Layers)
  $Texts = @($Layers | Where-Object { $_.IsText -and $_.Name })
  $MarketHeading = $Texts | Where-Object { $_.Name -like "*해외시장뉴스*" } | Select-Object -First 1
  $BusinessHeading = $Texts | Where-Object { $_.Name -eq "KOTRA 사업안내" } | Select-Object -First 1
  if ($null -eq $MarketHeading) {
    if ($BestEffort) { $Warnings.Add("PSD market-news heading was not found.") } else { throw "PSD market-news heading was not found." }
  }
  if ($null -eq $BusinessHeading) {
    if ($BestEffort) { $Warnings.Add("PSD business heading was not found.") } else { throw "PSD business heading was not found." }
  }
  $Headline = ""
  $MarketItems = @()
  if ($null -ne $MarketHeading) {
    $SectionNames = @("KOTRA 사업안내", "아카데미 교육·연수 과정", "경제외교 소식")
    $Following = @($Texts | Where-Object { $SectionNames -contains $_.Name -and $_.Top -gt $MarketHeading.Top } | ForEach-Object Top)
    $MarketBottom = if ($Following.Count -gt 0) { ($Following | Measure-Object -Minimum).Minimum } else { ($Texts | Measure-Object Bottom -Maximum).Maximum + 1 }
    $Zone = @($Texts | Where-Object { $_.Top -gt $MarketHeading.Top -and $_.Top -lt $MarketBottom })
    $Locations = @($Zone | Where-Object { $null -ne (Split-MarketLocation $_.Name) })
    if ($Locations.Count -ne 6) {
      $Message = "PSD has $($Locations.Count) market office layers instead of 6."
      if ($BestEffort) { $Warnings.Add($Message) } else { throw $Message }
    }
    $Used = @{}
    $LocationIndexes = @{}; foreach ($Location in $Locations) { $LocationIndexes[$Location.Index] = $true }
    foreach ($Location in $Locations) {
      $Candidates = @($Zone | Where-Object {
        -not $Used.ContainsKey($_.Index) -and -not $LocationIndexes.ContainsKey($_.Index) -and
        $_.Name -notlike "*해외시장뉴스*" -and $_.Name.Trim([char[]]"…·. ") -and
        $_.Top - $Location.Top -ge -5 -and $_.Top - $Location.Top -le 80 -and
        [Math]::Abs($_.Left - $Location.Left) -le 180
      })
      if ($Candidates.Count -eq 0) {
        $Message = "No article title matched PSD layer '$($Location.Name)'."
        if ($BestEffort) { $Warnings.Add($Message); continue } else { throw $Message }
      }
      $Title = $Candidates | Sort-Object @{ Expression = { [Math]::Abs($_.Top - $Location.Top) } }, @{ Expression = { [Math]::Abs($_.Left - $Location.Left) } } | Select-Object -First 1
      $Used[$Title.Index] = $true
      $Parts = Split-MarketLocation $Location.Name
      $MarketItems += [pscustomobject]@{ Country = $Parts.Country; Office = $Parts.Office; Title = Get-CleanText $Title.Name; Top = $Location.Top; Left = $Location.Left; MapSlot = 0 }
    }
    try { $MarketItems = @(Assign-MarketSlots $MarketItems) } catch { if ($BestEffort) { $Warnings.Add($_.Exception.Message) } else { throw } }
    $FirstTop = if ($Locations.Count -gt 0) { ($Locations | Measure-Object Top -Minimum).Minimum } else { $null }
    $HeadlineCandidates = @($Zone | Where-Object { $null -ne $FirstTop -and $_.Top -lt $FirstTop -and $_.Name -notlike "*해외시장뉴스*" -and $null -eq (Split-MarketLocation $_.Name) -and $_.Name.Trim([char[]]"…·. ") })
    if ($HeadlineCandidates.Count -gt 0) { $Headline = ($HeadlineCandidates | Sort-Object { $_.Name.Length } -Descending | Select-Object -First 1).Name }
    else { if ($BestEffort) { $Warnings.Add("PSD market-news headline was not found.") } else { throw "PSD market-news headline was not found." } }
  }
  $EconomyHeading = $Texts | Where-Object { $_.Name -eq "경제외교 소식" } | Select-Object -First 1
  $EconomyItems = @()
  if ($null -ne $EconomyHeading) {
    $EconomyLayers = @($Texts | Where-Object { $_.Top -gt $EconomyHeading.Top -and $_.Left -lt 560 -and ($_.Name.Contains("ㅣ") -or $_.Name.Contains("|")) } | Sort-Object Top | Select-Object -First 2)
    foreach ($Layer in $EconomyLayers) {
      $Parts = (New-Object Text.RegularExpressions.Regex("[ㅣ|]")).Split($Layer.Name, 2)
      $EconomyItems += [pscustomobject]@{ Label = Get-CleanText $Parts[0]; Title = Get-CleanText $Parts[1]; MapSlot = if ($Layer.Top - $EconomyHeading.Top -lt 60) { 1 } else { 2 } }
    }
  }
  if ($EconomyItems.Count -ne 2) {
    $Message = "PSD has $($EconomyItems.Count) economy-diplomacy items instead of 2."
    if ($BestEffort) { $Warnings.Add($Message) } else { throw $Message }
  }
  $EconomyBanner = ($Texts | Where-Object { $_.Name -like "*Help Desk*" } | Select-Object -First 1).Name
  $PsdDate = $null
  foreach ($Layer in $Texts) {
    $Match = [regex]::Match($Layer.Name, "^(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일$")
    if ($Match.Success) {
      $PsdDate = New-Object DateTime -ArgumentList ([int]$Match.Groups[1].Value), ([int]$Match.Groups[2].Value), ([int]$Match.Groups[3].Value)
      break
    }
  }
  return [pscustomobject]@{ Headline = $Headline; Market = $MarketItems; Economy = $EconomyItems; EconomyBanner = $EconomyBanner; Date = $PsdDate }
}

function Find-Container {
  param([string]$Html, [string]$ClassName, [int]$Start = 0)
  $EscapedClass = [regex]::Escape($ClassName)
  $Pattern = "<div\b(?=[^>]*\bclass\s*=\s*([`"'])[^`"']*\b$EscapedClass\b[^`"']*\1)[^>]*>"
  $OpenRegex = New-Object Text.RegularExpressions.Regex($Pattern, [Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [Text.RegularExpressions.RegexOptions]::Singleline)
  $Open = $OpenRegex.Match($Html, $Start)
  if (-not $Open.Success) { throw "HTML class '$ClassName' container not found." }
  $TagRegex = New-Object Text.RegularExpressions.Regex("<div\b[^>]*>|</div\s*>", [Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [Text.RegularExpressions.RegexOptions]::Singleline)
  $Depth = 1
  foreach ($Tag in $TagRegex.Matches($Html, $Open.Index + $Open.Length)) {
    if ($Tag.Value.ToLowerInvariant().StartsWith("</div")) {
      $Depth--
      if ($Depth -eq 0) { return [pscustomobject]@{ Start = $Open.Index; OpenEnd = $Open.Index + $Open.Length; CloseStart = $Tag.Index; End = $Tag.Index + $Tag.Length } }
    } else { $Depth++ }
  }
  throw "HTML class '$ClassName' closing div not found."
}

function Set-ContainerInner {
  param([string]$Html, [string]$ClassName, [string]$Inner, [int]$Start = 0)
  $Span = Find-Container $Html $ClassName $Start
  return $Html.Substring(0, $Span.OpenEnd) + $Inner + $Html.Substring($Span.CloseStart)
}

function Set-NamedAnchor {
  param([string]$Html, [string]$ClassName, [string]$Href, [string]$ImageAlt)
  $EscapedClass = [regex]::Escape($ClassName)
  $Pattern = "<a\b(?=[^>]*\bclass\s*=\s*([`"'])[^`"']*\b$EscapedClass\b[^`"']*\1)[^>]*>.*?</a\s*>"
  $Regex = New-Object Text.RegularExpressions.Regex($Pattern, [Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [Text.RegularExpressions.RegexOptions]::Singleline)
  $Match = $Regex.Match($Html)
  if (-not $Match.Success) { throw "HTML anchor class '$ClassName' not found." }
  $Anchor = $Match.Value
  $OpenEnd = $Anchor.IndexOf(">") + 1
  $Anchor = (Set-AttrValue $Anchor.Substring(0, $OpenEnd) "href" $Href) + $Anchor.Substring($OpenEnd)
  $ImageMatch = [regex]::Match($Anchor, $ImagePattern, "IgnoreCase,Singleline")
  if ($ImageMatch.Success) {
    $Image = Set-AttrValue $ImageMatch.Value "alt" $ImageAlt
    $Anchor = $Anchor.Substring(0, $ImageMatch.Index) + $Image + $Anchor.Substring($ImageMatch.Index + $ImageMatch.Length)
  }
  return $Html.Substring(0, $Match.Index) + $Anchor + $Html.Substring($Match.Index + $Match.Length)
}

function Set-HeaderDate {
  param([string]$Html, [DateTime]$Date)
  $Regex = New-Object Text.RegularExpressions.Regex('(<div\s+class="lt">\s*<h1>kotra 뉴스레터</h1>\s*<p>).*?(</p>)', [Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [Text.RegularExpressions.RegexOptions]::Singleline)
  if (-not $Regex.IsMatch($Html)) { throw "Newsletter header date not found in the template." }
  $Text = "$($Date.Year)년 $($Date.Month)월 $($Date.Day)일"
  return $Regex.Replace($Html, [Text.RegularExpressions.MatchEvaluator]{ param($m) $m.Groups[1].Value + $Text + $m.Groups[2].Value }, 1)
}

function Set-MarketItems {
  param([string]$Html, [object[]]$Items, [hashtable]$Links)
  $Span = Find-Container $Html "newsBoxes"
  $Inner = $Html.Substring($Span.OpenEnd, $Span.CloseStart - $Span.OpenEnd)
  $Anchors = @([regex]::Matches($Inner, $AnchorPattern, "IgnoreCase,Singleline") | Where-Object { $_.Value -match "\bclass\s*=\s*([`"'])[^`"']*\bnewsBox\b" })
  if ($Anchors.Count -ne 6) { throw "Template has $($Anchors.Count) market boxes instead of 6." }
  $BySlot = @{}; foreach ($Item in $Items) { if ($Item.MapSlot) { $BySlot[[int]$Item.MapSlot] = $Item } }
  $Replacements = @()
  for ($Position = 0; $Position -lt $MarketLinkSlots.Count; $Position++) {
    $Slot = $MarketLinkSlots[$Position]
    if (-not $BySlot.ContainsKey($Slot) -or -not $Links["리스트$Slot"]) { continue }
    $Item = $BySlot[$Slot]
    $Box = '<a href="' + (ConvertTo-HtmlText $Links["리스트$Slot"]) + '" class="newsBox" target="_blank">' + "`n" +
      '                    <div class="local">' + "`n" +
      '                        <span>' + (ConvertTo-HtmlText $Item.Country) + '</span>' + "`n" +
      '                        <span>' + (ConvertTo-HtmlText $Item.Office) + '</span>' + "`n" +
      '                    </div>' + "`n" +
      '                    <p>' + (ConvertTo-HtmlText $Item.Title) + '</p>' + "`n" +
      '                </a>'
    $Replacements += [pscustomobject]@{ Match = $Anchors[$Position]; Text = $Box }
  }
  foreach ($Replacement in @($Replacements | Sort-Object { $_.Match.Index } -Descending)) {
    $Match = $Replacement.Match
    $Inner = $Inner.Substring(0, $Match.Index) + $Replacement.Text + $Inner.Substring($Match.Index + $Match.Length)
  }
  return [pscustomobject]@{ Html = $Html.Substring(0, $Span.OpenEnd) + $Inner + $Html.Substring($Span.CloseStart); Count = $Replacements.Count }
}

function Set-ProjectSection {
  param([string]$Html, [string]$Heading, [object[]]$Records, [string]$PeriodLabel)
  $HeadingText = "<h2>$Heading</h2>"
  $HeadingPosition = $Html.IndexOf($HeadingText)
  if ($HeadingPosition -lt 0) { throw "Template heading not found: $Heading" }
  $Builder = New-Object Text.StringBuilder
  foreach ($Record in $Records) {
    [void]$Builder.Append("`n                <a href=`"").Append((ConvertTo-HtmlText $Record.url)).Append('" class="projectBox">').Append("`n")
    [void]$Builder.Append("                    <ul>`n                        <li class=`"period`">`n                            <span>").Append($PeriodLabel).Append("</span>`n                            <p>").Append((ConvertTo-HtmlText $Record.period)).Append("</p>`n                        </li>`n")
    [void]$Builder.Append("                        <li class=`"deadline`">`n                            <span>신청마감</span>`n                            <p>").Append((ConvertTo-HtmlText $Record.deadline)).Append("</p>`n                        </li>`n                    </ul>`n")
    [void]$Builder.Append("                    <p>").Append((ConvertTo-HtmlText $Record.title)).Append("</p>`n                </a>")
  }
  [void]$Builder.Append("`n            ")
  return Set-ContainerInner $Html "projectBoxes" $Builder.ToString() $HeadingPosition
}

function Set-Banners {
  param([string]$Html, [hashtable]$Records, [hashtable]$Links)
  $Span = Find-Container $Html "bannerCont"
  $Inner = $Html.Substring($Span.OpenEnd, $Span.CloseStart - $Span.OpenEnd)
  $Anchors = @([regex]::Matches($Inner, $AnchorPattern, "IgnoreCase,Singleline"))
  if ($Anchors.Count -ne 3) { throw "Template has $($Anchors.Count) banners instead of 3." }
  $Texts = @()
  for ($Index = 0; $Index -lt 3; $Index++) {
    $Slot = $Index + 1
    $Anchor = $Anchors[$Index].Value
    $Record = $Records["banner|$Slot"]
    if ($Slot -eq 2) { $Href = if ($Links["홍보배너2"]) { $Links["홍보배너2"] } else { Get-AttrValue $Anchor "href" }; $Alt = "글로벌 방위산업 트렌드" }
    elseif ($null -ne $Record) { $Href = $Record.url; $Alt = $Record.title }
    else { $Href = Get-AttrValue $Anchor "href"; $Img = [regex]::Match($Anchor, $ImagePattern, "IgnoreCase,Singleline"); $Alt = if ($Img.Success) { Get-AttrValue $Img.Value "alt" } else { "" } }
    $OpenEnd = $Anchor.IndexOf(">") + 1
    $Anchor = (Set-AttrValue $Anchor.Substring(0, $OpenEnd) "href" $Href) + $Anchor.Substring($OpenEnd)
    $Img = [regex]::Match($Anchor, $ImagePattern, "IgnoreCase,Singleline")
    if ($Img.Success) { $NewImg = Set-AttrValue $Img.Value "alt" $Alt; $Anchor = $Anchor.Substring(0, $Img.Index) + $NewImg + $Anchor.Substring($Img.Index + $Img.Length) }
    $Texts += $Anchor
  }
  for ($Index = $Anchors.Count - 1; $Index -ge 0; $Index--) { $Inner = $Inner.Substring(0, $Anchors[$Index].Index) + $Texts[$Index] + $Inner.Substring($Anchors[$Index].Index + $Anchors[$Index].Length) }
  return $Html.Substring(0, $Span.OpenEnd) + $Inner + $Html.Substring($Span.CloseStart)
}

function Set-PrimaryOrder {
  param([string]$Html, [string]$Order)
  $Spans = @{}
  $Start = 0
  while ($true) {
    try { $Span = Find-Container $Html "nlCont" $Start } catch { break }
    $Block = $Html.Substring($Span.Start, $Span.End - $Span.Start)
    if ($Block.Contains("<h2>KOTRA 해외시장뉴스</h2>")) { $Spans.market = $Span }
    if ($Block.Contains("<h2>KOTRA 사업안내</h2>")) { $Spans.business = $Span }
    $Start = $Span.End
  }
  if (-not $Spans.market -or -not $Spans.business) { throw "Template market/business blocks not found." }
  $IsBusinessFirst = $Spans.business.Start -lt $Spans.market.Start
  if ($IsBusinessFirst -eq ($Order -eq "business_first")) { return $Html }
  $First = if ($Spans.market.Start -lt $Spans.business.Start) { $Spans.market } else { $Spans.business }
  $Second = if ($Spans.market.Start -lt $Spans.business.Start) { $Spans.business } else { $Spans.market }
  $Between = $Html.Substring($First.End, $Second.Start - $First.End)
  if ($Between -match "<div\b[^>]*\bclass\s*=\s*([`"'])[^`"']*\bnlCont\b") { throw "Another newsletter section prevents safe reordering." }
  return $Html.Substring(0, $First.Start) + $Html.Substring($Second.Start, $Second.End - $Second.Start) + $Between + $Html.Substring($First.Start, $First.End - $First.Start) + $Html.Substring($Second.End)
}

function Get-ImageSources {
  param([string]$Html)
  return @([regex]::Matches($Html, $ImagePattern, "IgnoreCase,Singleline") | ForEach-Object { Get-AttrValue $_.Value "src" })
}

function Test-StringArraysEqual {
  param([string[]]$Left, [string[]]$Right)
  if ($Left.Count -ne $Right.Count) { return $false }
  for ($Index = 0; $Index -lt $Left.Count; $Index++) { if ($Left[$Index] -ne $Right[$Index]) { return $false } }
  return $true
}

function Invoke-SafeUpdate {
  param([string]$Label, [scriptblock]$Operation, [ref]$Html)
  try { $Html.Value = & $Operation $Html.Value } catch { if ($BestEffort) { $Warnings.Add("$Label`: $($_.Exception.Message)") } else { throw } }
}

try {
  $TemplateBytes = [IO.File]::ReadAllBytes((Resolve-Path -LiteralPath $Template).Path)
  $HadBom = $TemplateBytes.Length -ge 3 -and $TemplateBytes[0] -eq 0xEF -and $TemplateBytes[1] -eq 0xBB -and $TemplateBytes[2] -eq 0xBF
  $TemplateHtml = [Text.Encoding]::UTF8.GetString($TemplateBytes)
  if ($TemplateHtml.Length -gt 0 -and $TemplateHtml[0] -eq [char]0xFEFF) { $TemplateHtml = $TemplateHtml.Substring(1) }
  $UseCrLf = $TemplateHtml.Contains("`r`n")

  try { $CsvData = Read-CsvRecords (Resolve-Path -LiteralPath $Csv).Path } catch { if (-not $BestEffort) { throw }; $Warnings.Add("CSV: $($_.Exception.Message)"); $CsvData = [pscustomobject]@{ Records = @{}; Date = $null } }
  try { $Links = Read-MapLinks (Resolve-Path -LiteralPath $MapHtml).Path } catch { if (-not $BestEffort) { throw }; $Warnings.Add("Image map: $($_.Exception.Message)"); $Links = @{} }
  try { $SectionOrder = Get-PrimarySectionOrder (Resolve-Path -LiteralPath $MapHtml).Path } catch { if (-not $BestEffort) { throw }; $Warnings.Add("Section order: $($_.Exception.Message)"); $SectionOrder = if ($TemplateHtml.IndexOf("<h2>KOTRA 사업안내</h2>") -lt $TemplateHtml.IndexOf("<h2>KOTRA 해외시장뉴스</h2>")) { "business_first" } else { "market_first" } }
  try { $PsdData = Extract-PsdContent @(Read-PsdLayers (Resolve-Path -LiteralPath $Psd).Path) } catch { if (-not $BestEffort) { throw }; $Warnings.Add("PSD: $($_.Exception.Message)"); $PsdData = [pscustomobject]@{ Headline = ""; Market = @(); Economy = @(); EconomyBanner = ""; Date = $null } }

  $ExplicitDate = $null
  if ($IssueDate) {
    [DateTime]$Parsed = [DateTime]::MinValue
    if ([DateTime]::TryParseExact($IssueDate, "yyyy-MM-dd", [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$Parsed)) { $ExplicitDate = $Parsed.Date }
    elseif ($BestEffort) { $Warnings.Add("Invalid IssueDate: $IssueDate") } else { throw "Invalid IssueDate: $IssueDate" }
  }
  $FinalDate = if ($null -ne $ExplicitDate) { $ExplicitDate } elseif ($null -ne $CsvData.Date) { $CsvData.Date } else { $PsdData.Date }
  $Result = $TemplateHtml
  if ($null -ne $FinalDate) { Invoke-SafeUpdate "Issue date" { param($h) Set-HeaderDate $h $FinalDate } ([ref]$Result) } else { if ($BestEffort) { $Warnings.Add("Issue date was not found; template value was kept.") } else { throw "Issue date was not found." } }
  if ($PsdData.Headline -and $Links["헤드라인"]) { Invoke-SafeUpdate "Headline" { param($h) Set-NamedAnchor $h "titleNews" $Links["헤드라인"] $PsdData.Headline } ([ref]$Result) } else { $Warnings.Add("Headline text or URL is missing; template value was kept.") }
  if (@($PsdData.Market).Count -gt 0) {
    try { $Changed = Set-MarketItems $Result @($PsdData.Market) $Links; $Result = $Changed.Html; if ($Changed.Count -lt @($PsdData.Market).Count) { $Warnings.Add("Only $($Changed.Count) market items with URLs were filled.") } } catch { if ($BestEffort) { $Warnings.Add("Market news: $($_.Exception.Message)") } else { throw } }
  } else { $Warnings.Add("Market-news PSD text is missing; template value was kept.") }
  $Business = @(1..8 | ForEach-Object { $CsvData.Records["business|$_"] } | Where-Object { $null -ne $_ })
  $Academy = @(1..4 | ForEach-Object { $CsvData.Records["academy|$_"] } | Where-Object { $null -ne $_ })
  if ($Business.Count -gt 0) { Invoke-SafeUpdate "Business" { param($h) Set-ProjectSection $h "KOTRA 사업안내" $Business "사업기간" } ([ref]$Result) } else { $Warnings.Add("Business CSV data is missing; template value was kept.") }
  if ($Academy.Count -gt 0) { Invoke-SafeUpdate "Academy" { param($h) Set-ProjectSection $h "아카데미 교육·연수 과정" $Academy "교육기간" } ([ref]$Result) } else { $Warnings.Add("Academy CSV data is missing; template value was kept.") }
  if (@($CsvData.Records.Keys | Where-Object { $_ -like "banner|*" }).Count -gt 0 -or $Links["홍보배너2"]) { Invoke-SafeUpdate "Banners" { param($h) Set-Banners $h $CsvData.Records $Links } ([ref]$Result) }
  $EconomyItems = @($PsdData.Economy | Where-Object { $Links["경제외교$($_.MapSlot)"] })
  if ($EconomyItems.Count -gt 0) {
    $Builder = New-Object Text.StringBuilder
    foreach ($Item in $EconomyItems) { [void]$Builder.Append("`n                    <a href=`"").Append((ConvertTo-HtmlText $Links["경제외교$($Item.MapSlot)"])).Append('" class="exampleBox">').Append("`n                        <span>").Append((ConvertTo-HtmlText $Item.Label)).Append("</span>`n                        <p>").Append((ConvertTo-HtmlText $Item.Title)).Append("</p>`n                    </a>") }
    [void]$Builder.Append("`n                ")
    $EconomyInner = $Builder.ToString()
    Invoke-SafeUpdate "Economy" { param($h) Set-ContainerInner $h "exampleBoxes" $EconomyInner } ([ref]$Result)
  } else { $Warnings.Add("Economy text or URLs are missing; template value was kept.") }
  if ($Links["경제외교 우측배너"]) {
    $EconomyAlt = if ($PsdData.EconomyBanner) { $PsdData.EconomyBanner } else { "경제외교 소식" }
    Invoke-SafeUpdate "Economy banner" { param($h) Set-NamedAnchor $h "exampleImg" $Links["경제외교 우측배너"] $EconomyAlt } ([ref]$Result)
  }
  else { $Warnings.Add("Economy banner URL is missing; template value was kept.") }
  Invoke-SafeUpdate "Section order" { param($h) Set-PrimaryOrder $h $SectionOrder } ([ref]$Result)

  $Result = $Result.Replace("`r`n", "`n").Replace("`r", "`n")
  if ($UseCrLf) { $Result = $Result.Replace("`n", "`r`n") }
  if (-not (Test-StringArraysEqual @(Get-ImageSources $TemplateHtml) @(Get-ImageSources $Result))) { throw "An image src changed; output was not saved." }
  $Parent = Split-Path -Parent $Output
  if ($Parent) { [IO.Directory]::CreateDirectory($Parent) | Out-Null }
  $Encoding = New-Object Text.UTF8Encoding($HadBom)
  [IO.File]::WriteAllText($Output, $Result, $Encoding)
  Write-Host "Mutu24 HTML: $Output"
  Write-Host ("Filled: headline {0}, market {1}, business {2}, academy {3}, banners {4}, economy {5}" -f $(if ($PsdData.Headline) { 1 } else { 0 }), @($PsdData.Market).Count, $Business.Count, $Academy.Count, (@($CsvData.Records.Keys | Where-Object { $_ -like "banner|*" }).Count + 1), @($PsdData.Economy).Count)
  Write-Host "Image src: unchanged"
  Write-Host ("Top order: " + $(if ($SectionOrder -eq "business_first") { "business -> market news" } else { "market news -> business" }))
  foreach ($Warning in $Warnings) { Write-Warning $Warning }
  if ($Warnings.Count -gt 0) { exit 2 }
  exit 0
} catch {
  Write-Error $_.Exception.Message
  exit 1
}
