[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)][string]$Xlsx,
  [Parameter(Mandatory = $true)][string]$OutputDir,
  [string]$AssetsDir,
  [string]$NewsletterDate,
  [switch]$BestEffort
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$MainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
$RelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
$PkgRelNs = "http://schemas.openxmlformats.org/package/2006/relationships"
$Warnings = New-Object System.Collections.Generic.List[string]
$Capacities = @{ business = 8; banner = 3; academy = 4 }
$WeekdaysKo = @("일", "월", "화", "수", "목", "금", "토")

function Get-CleanText {
  param($Value)
  if ($null -eq $Value) { return "" }
  return ([regex]::Replace(([string]$Value).Replace([char]0x00A0, " ").Replace("`r", " ").Replace("`n", " "), "\s+", " ")).Trim()
}

function Get-ZipText {
  param([System.IO.Compression.ZipArchive]$Package, [string]$Name)
  $Entry = $Package.GetEntry($Name)
  if ($null -eq $Entry) { throw "Excel package entry not found: $Name" }
  $Reader = New-Object System.IO.StreamReader($Entry.Open(), [Text.Encoding]::UTF8, $true)
  try { return $Reader.ReadToEnd() } finally { $Reader.Dispose() }
}

function Get-ZipBytes {
  param([System.IO.Compression.ZipArchive]$Package, [string]$Name)
  $Entry = $Package.GetEntry($Name)
  if ($null -eq $Entry) { throw "Excel package entry not found: $Name" }
  $Stream = $Entry.Open()
  $Memory = New-Object System.IO.MemoryStream
  try {
    $Stream.CopyTo($Memory)
    return $Memory.ToArray()
  } finally {
    $Memory.Dispose()
    $Stream.Dispose()
  }
}

function Get-ZipXml {
  param([System.IO.Compression.ZipArchive]$Package, [string]$Name)
  [xml]$Xml = Get-ZipText $Package $Name
  return $Xml
}

function Resolve-Part {
  param([string]$BasePart, [string]$Target)
  if ($Target.StartsWith("/")) { return $Target.TrimStart("/") }
  $Slash = $BasePart.LastIndexOf("/")
  $BaseDir = if ($Slash -ge 0) { $BasePart.Substring(0, $Slash + 1) } else { "" }
  $Stack = New-Object System.Collections.Generic.List[string]
  foreach ($Part in ($BaseDir + $Target).Split("/")) {
    if (-not $Part -or $Part -eq ".") { continue }
    if ($Part -eq "..") {
      if ($Stack.Count -gt 0) { $Stack.RemoveAt($Stack.Count - 1) }
    } else {
      $Stack.Add($Part)
    }
  }
  return [string]::Join("/", $Stack)
}

function Get-RelsPart {
  param([string]$Part)
  $Slash = $Part.LastIndexOf("/")
  if ($Slash -lt 0) { return "_rels/$Part.rels" }
  return $Part.Substring(0, $Slash) + "/_rels/" + $Part.Substring($Slash + 1) + ".rels"
}

function Get-RelationshipMap {
  param([xml]$Xml)
  $Map = @{}
  foreach ($Node in $Xml.SelectNodes("//*[local-name()='Relationship']")) {
    $Map[[string]$Node.Id] = [string]$Node.Target
  }
  return $Map
}

function Get-ColumnNumber {
  param([string]$CellReference)
  $Match = [regex]::Match($CellReference, "^[A-Z]+")
  if (-not $Match.Success) { throw "Invalid cell reference: $CellReference" }
  $Value = 0
  foreach ($Char in $Match.Value.ToCharArray()) { $Value = $Value * 26 + ([int]$Char - 64) }
  return $Value
}

function Get-SharedStrings {
  param([System.IO.Compression.ZipArchive]$Package)
  if ($null -eq $Package.GetEntry("xl/sharedStrings.xml")) { return @() }
  $Xml = Get-ZipXml $Package "xl/sharedStrings.xml"
  $Result = @()
  foreach ($Item in $Xml.SelectNodes("//*[local-name()='si']")) {
    $Text = ""
    foreach ($Node in $Item.SelectNodes(".//*[local-name()='t']")) { $Text += [string]$Node.InnerText }
    $Result += $Text
  }
  return $Result
}

function Get-TargetSheet {
  param([System.IO.Compression.ZipArchive]$Package)
  $WorkbookPart = "xl/workbook.xml"
  $Workbook = Get-ZipXml $Package $WorkbookPart
  $Sheets = @($Workbook.SelectNodes("//*[local-name()='sheets']/*[local-name()='sheet']"))
  if ($Sheets.Count -eq 0) { throw "The workbook has no worksheets." }
  $Selected = $null
  foreach ($Wanted in @("최종", "최종 업로드")) {
    $Selected = $Sheets | Where-Object { [string]$_.name -eq $Wanted } | Select-Object -First 1
    if ($null -ne $Selected) { break }
  }
  if ($null -eq $Selected) { $Selected = $Sheets[0] }
  $Relationships = Get-RelationshipMap (Get-ZipXml $Package (Get-RelsPart $WorkbookPart))
  $RelId = $Selected.GetAttribute("id", $RelNs)
  if (-not $RelId -or -not $Relationships.ContainsKey($RelId)) { throw "Could not resolve the target worksheet." }
  return [pscustomobject]@{
    Name = [string]$Selected.name
    Part = Resolve-Part $WorkbookPart $Relationships[$RelId]
  }
}

function Read-TargetSheet {
  param([System.IO.Compression.ZipArchive]$Package)
  $Shared = @(Get-SharedStrings $Package)
  $Target = Get-TargetSheet $Package
  $Xml = Get-ZipXml $Package $Target.Part
  $Rows = @{}
  foreach ($Cell in $Xml.SelectNodes("//*[local-name()='sheetData']/*[local-name()='row']/*[local-name()='c']")) {
    $Reference = [string]$Cell.r
    $RowMatch = [regex]::Match($Reference, "\d+$")
    if (-not $RowMatch.Success) { continue }
    $RowNo = [int]$RowMatch.Value
    $ColNo = Get-ColumnNumber $Reference
    $CellType = [string]$Cell.t
    $ValueNode = $Cell.SelectSingleNode("./*[local-name()='v']")
    $Value = $null
    if ($CellType -eq "inlineStr") {
      $Value = ""
      foreach ($TextNode in $Cell.SelectNodes(".//*[local-name()='t']")) { $Value += [string]$TextNode.InnerText }
    } elseif ($null -ne $ValueNode) {
      $Raw = [string]$ValueNode.InnerText
      if ($CellType -eq "s") {
        $Value = $Shared[[int]$Raw]
      } elseif ($CellType -eq "b") {
        $Value = ($Raw -eq "1")
      } elseif ($CellType -eq "str" -or $CellType -eq "e") {
        $Value = $Raw
      } else {
        [double]$Number = 0
        if ([double]::TryParse($Raw, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$Number)) {
          if ($Number -eq [Math]::Truncate($Number)) { $Value = [long]$Number } else { $Value = $Number }
        } else { $Value = $Raw }
      }
    }
    if (-not $Rows.ContainsKey($RowNo)) { $Rows[$RowNo] = @{} }
    $Rows[$RowNo][$ColNo] = $Value
  }
  return [pscustomobject]@{ Name = $Target.Name; Part = $Target.Part; Rows = $Rows }
}

function Get-RowValue {
  param([hashtable]$Rows, [int]$Row, [int]$Column)
  if ($Rows.ContainsKey($Row) -and $Rows[$Row].ContainsKey($Column)) { return $Rows[$Row][$Column] }
  return $null
}

function ConvertFrom-ExcelDate {
  param($Value)
  if ($Value -is [DateTime]) { return ([DateTime]$Value).Date }
  if ($Value -is [int] -or $Value -is [long] -or $Value -is [double] -or $Value -is [decimal]) {
    return (Get-Date "1899-12-30").AddDays([double]$Value).Date
  }
  $Text = Get-CleanText $Value
  $Match = [regex]::Match($Text, "^(\d{4})[-./](\d{1,2})[-./](\d{1,2})")
  if ($Match.Success) {
    return New-Object DateTime -ArgumentList ([int]$Match.Groups[1].Value), ([int]$Match.Groups[2].Value), ([int]$Match.Groups[3].Value)
  }
  return $null
}

function Format-KoreanDate {
  param([DateTime]$Value, [bool]$IncludeYear = $false)
  $Prefix = if ($IncludeYear) { "{0:00}." -f ($Value.Year % 100) } else { "" }
  return "{0}{1}.{2}.({3})" -f $Prefix, $Value.Month, $Value.Day, $WeekdaysKo[[int]$Value.DayOfWeek]
}

function Format-Period {
  param($StartValue, $EndValue, [DateTime]$IssueDate)
  $Start = ConvertFrom-ExcelDate $StartValue
  $End = ConvertFrom-ExcelDate $EndValue
  if ($null -eq $Start) { return Get-CleanText $StartValue }
  $StartText = Format-KoreanDate $Start ($null -ne $IssueDate -and $Start.Year -ne $IssueDate.Year)
  if ($null -eq $End -or $End -eq $Start) { return $StartText }
  $ShowEndYear = $End.Year -ne $Start.Year -or ($null -ne $IssueDate -and $End.Year -ne $IssueDate.Year)
  return "$StartText  -  $(Format-KoreanDate $End $ShowEndYear)"
}

function Normalize-Url {
  param($Value)
  $Text = Get-CleanText $Value
  $Match = [regex]::Match($Text, "https?://[^\s\]\[()<>]+", [Text.RegularExpressions.RegexOptions]::IgnoreCase)
  if ($Match.Success) { $Text = $Match.Value.TrimEnd(".", ",", ";") }
  return [regex]::Replace($Text, "^(https?://)(kotra\.or\.kr)(?=[:/?#]|$)", '$1www.$2', [Text.RegularExpressions.RegexOptions]::IgnoreCase)
}

function Get-NumericSlot {
  param($Value)
  if ($Value -is [int] -or $Value -is [long] -or $Value -is [double]) { return [int]$Value }
  $Match = [regex]::Match((Get-CleanText $Value), "\d+")
  if ($Match.Success) { return [int]$Match.Value }
  return $null
}

function Get-Sections {
  param([hashtable]$Rows)
  $Labels = @()
  foreach ($RowNo in @($Rows.Keys | Sort-Object)) {
    $Compact = (Get-CleanText (Get-RowValue $Rows $RowNo 1)).Replace(" ", "")
    $Section = switch ($Compact) {
      "사업안내형" { "business" }
      "사업안내" { "business" }
      "배너형" { "banner" }
      "배너" { "banner" }
      "아카데미" { "academy" }
      default { $null }
    }
    if ($Section) { $Labels += [pscustomobject]@{ Row = [int]$RowNo; Section = $Section } }
  }
  $Sections = @{}
  for ($Index = 0; $Index -lt $Labels.Count; $Index++) {
    $Label = $Labels[$Index]
    $End = if ($Index + 1 -lt $Labels.Count) { $Labels[$Index + 1].Row - 1 } else { [int](@($Rows.Keys | Measure-Object -Maximum).Maximum) }
    $Header = 0
    for ($RowNo = $Label.Row + 1; $RowNo -le $End; $RowNo++) {
      if ((Get-CleanText (Get-RowValue $Rows $RowNo 1)) -eq "연번") { $Header = $RowNo; break }
    }
    if ($Header -eq 0) { throw "Could not find the serial-number header in section $($Label.Section)." }
    $Sections[$Label.Section] = @($Header + 1, $End)
  }
  foreach ($Required in @("business", "banner", "academy")) {
    if (-not $Sections.ContainsKey($Required)) { throw "Required section not found in the final worksheet: $Required" }
  }
  return $Sections
}

function Add-RecordOrWarning {
  param([System.Collections.ArrayList]$Records, [string]$Section, [int]$Slot, [string]$Title, [string]$Period, [string]$Deadline, [string]$Url)
  if ($Slot -lt 1 -or $Slot -gt $Capacities[$Section]) {
    $Message = "$Section slot $Slot exceeds the supported capacity $($Capacities[$Section])."
    if ($BestEffort) { $Warnings.Add($Message); return }
    throw $Message
  }
  [void]$Records.Add([ordered]@{
    section = $Section; slot = [string]$Slot; title = $Title; period = $Period
    deadline = $Deadline; url = $Url; image_path = ""; newsletter_date = ""
  })
}

function Get-Records {
  param([hashtable]$Rows, [DateTime]$IssueDate)
  $Sections = Get-Sections $Rows
  $Records = New-Object System.Collections.ArrayList

  $Range = $Sections.business
  $Header = $Range[0] - 1
  $Compact = (Get-CleanText (Get-RowValue $Rows $Header 3)).Replace(" ", "") -eq "사업기간" -and
    @("신청마감", "접수마감") -contains (Get-CleanText (Get-RowValue $Rows $Header 4)).Replace(" ", "") -and
    (Get-CleanText (Get-RowValue $Rows $Header 5)).ToUpperInvariant() -eq "URL"
  for ($RowNo = $Range[0]; $RowNo -le $Range[1]; $RowNo++) {
    $Slot = Get-NumericSlot (Get-RowValue $Rows $RowNo 1)
    $Title = Get-CleanText (Get-RowValue $Rows $RowNo 2)
    if ($null -eq $Slot -or -not $Title -or $Title.StartsWith("입력 안내:")) { continue }
    if ($Compact) {
      $Period = Get-CleanText (Get-RowValue $Rows $RowNo 3)
      $Deadline = Get-CleanText (Get-RowValue $Rows $RowNo 4)
      $Url = Normalize-Url (Get-RowValue $Rows $RowNo 5)
    } else {
      $Period = Format-Period (Get-RowValue $Rows $RowNo 3) (Get-RowValue $Rows $RowNo 4) $IssueDate
      $DeadlineDate = ConvertFrom-ExcelDate (Get-RowValue $Rows $RowNo 5)
      $Deadline = if ($null -ne $DeadlineDate) { Format-KoreanDate $DeadlineDate ($DeadlineDate.Year -ne $IssueDate.Year) } else { Get-CleanText (Get-RowValue $Rows $RowNo 9) }
      $Url = Normalize-Url (Get-RowValue $Rows $RowNo 6)
    }
    Add-RecordOrWarning $Records "business" $Slot $Title $Period $Deadline $Url
  }

  $Range = $Sections.academy
  $Header = $Range[0] - 1
  $Compact = (Get-CleanText (Get-RowValue $Rows $Header 3)).Replace(" ", "") -eq "교육기간" -and
    @("신청마감", "접수마감", "접수마감일") -contains (Get-CleanText (Get-RowValue $Rows $Header 4)).Replace(" ", "") -and
    (Get-CleanText (Get-RowValue $Rows $Header 5)).ToUpperInvariant() -eq "URL" -and
    -not (Get-CleanText (Get-RowValue $Rows $Header 6))
  for ($RowNo = $Range[0]; $RowNo -le $Range[1]; $RowNo++) {
    $Slot = Get-NumericSlot (Get-RowValue $Rows $RowNo 1)
    $Title = Get-CleanText (Get-RowValue $Rows $RowNo 2)
    if ($null -eq $Slot -or -not $Title -or $Title.StartsWith("입력 안내:")) { continue }
    if ($Compact) {
      $Period = Get-CleanText (Get-RowValue $Rows $RowNo 3)
      $Deadline = Get-CleanText (Get-RowValue $Rows $RowNo 4)
    } else {
      $Period = Get-CleanText (Get-RowValue $Rows $RowNo 6)
      if (-not $Period) { $Period = Get-CleanText (Get-RowValue $Rows $RowNo 3) }
      $Deadline = Get-CleanText (Get-RowValue $Rows $RowNo 7)
      if (-not $Deadline) { $Deadline = Get-CleanText (Get-RowValue $Rows $RowNo 4) }
    }
    Add-RecordOrWarning $Records "academy" $Slot $Title $Period $Deadline (Normalize-Url (Get-RowValue $Rows $RowNo 5))
  }

  $Range = $Sections.banner
  for ($RowNo = $Range[0]; $RowNo -le $Range[1]; $RowNo++) {
    $Slot = Get-NumericSlot (Get-RowValue $Rows $RowNo 1)
    $Title = Get-CleanText (Get-RowValue $Rows $RowNo 2)
    if ($null -eq $Slot -or -not $Title -or $Title.StartsWith("입력 안내:")) { continue }
    $Url = Normalize-Url (Get-RowValue $Rows $RowNo 4)
    if (-not $Url) { continue }
    Add-RecordOrWarning $Records "banner" $Slot $Title "" (Get-CleanText (Get-RowValue $Rows $RowNo 3)) $Url
  }
  return $Records
}

function Get-SheetImages {
  param([System.IO.Compression.ZipArchive]$Package, [string]$SheetPart)
  $SheetRelsPart = Get-RelsPart $SheetPart
  if ($null -eq $Package.GetEntry($SheetRelsPart)) { return @() }
  $SheetRelsXml = Get-ZipXml $Package $SheetRelsPart
  $DrawingTarget = $null
  foreach ($Rel in $SheetRelsXml.SelectNodes("//*[local-name()='Relationship']")) {
    if ([string]$Rel.Type -like "*/drawing") { $DrawingTarget = [string]$Rel.Target; break }
  }
  if (-not $DrawingTarget) { return @() }
  $DrawingPart = Resolve-Part $SheetPart $DrawingTarget
  $DrawingRels = Get-RelationshipMap (Get-ZipXml $Package (Get-RelsPart $DrawingPart))
  $Drawing = Get-ZipXml $Package $DrawingPart
  $Pictures = @()
  foreach ($Anchor in $Drawing.SelectNodes("//*[local-name()='twoCellAnchor' or local-name()='oneCellAnchor']")) {
    $From = $Anchor.SelectSingleNode("./*[local-name()='from']")
    $Blip = $Anchor.SelectSingleNode(".//*[local-name()='blip']")
    if ($null -eq $From -or $null -eq $Blip) { continue }
    $RelId = $Blip.GetAttribute("embed", $RelNs)
    if (-not $RelId -or -not $DrawingRels.ContainsKey($RelId)) { continue }
    $Col = [int]$From.SelectSingleNode("./*[local-name()='col']").InnerText
    $Row = [int]$From.SelectSingleNode("./*[local-name()='row']").InnerText
    $ColOffNode = $From.SelectSingleNode("./*[local-name()='colOff']")
    $RowOffNode = $From.SelectSingleNode("./*[local-name()='rowOff']")
    $ColOff = if ($null -ne $ColOffNode) { [long]$ColOffNode.InnerText } else { 0 }
    $RowOff = if ($null -ne $RowOffNode) { [long]$RowOffNode.InnerText } else { 0 }
    $Pictures += [pscustomobject]@{
      X = [long]$Col * 10000000 + $ColOff
      Y = [long]$Row * 10000000 + $RowOff
      Part = Resolve-Part $DrawingPart $DrawingRels[$RelId]
    }
  }
  $Result = @()
  foreach ($Picture in @($Pictures | Sort-Object X, Y)) {
    $Result += [pscustomobject]@{
      Extension = [IO.Path]::GetExtension($Picture.Part).ToLowerInvariant()
      Bytes = Get-ZipBytes $Package $Picture.Part
    }
  }
  return $Result
}

function ConvertTo-CsvField {
  param($Value)
  $Text = [string]$Value
  if ($Text.Contains('"')) { $Text = $Text.Replace('"', '""') }
  if ($Text.IndexOfAny(@([char]',', [char]'"', [char]"`r", [char]"`n")) -ge 0) { return '"' + $Text + '"' }
  return $Text
}

function Write-NewsletterCsv {
  param([string]$Path, [System.Collections.IEnumerable]$Records, [DateTime]$IssueDate)
  $Encoding = New-Object Text.UTF8Encoding($true)
  $Writer = New-Object IO.StreamWriter($Path, $false, $Encoding)
  try {
    $Writer.WriteLine("section,slot,title,period,deadline,url,image_path,newsletter_date")
    $Index = 0
    foreach ($Record in $Records) {
      $DateText = if ($Index -eq 0) { $IssueDate.ToString("yyyy-MM-dd") } else { "" }
      $Values = @($Record.section, $Record.slot, $Record.title, $Record.period, $Record.deadline, $Record.url, $Record.image_path, $DateText)
      $Writer.WriteLine([string]::Join(",", @($Values | ForEach-Object { ConvertTo-CsvField $_ })))
      $Index++
    }
  } finally { $Writer.Dispose() }
}

try {
  $XlsxPath = (Resolve-Path -LiteralPath $Xlsx).Path
  $NameMatch = [regex]::Match([IO.Path]::GetFileName($XlsxPath), "^newsletter_DB_(20\d{6})\.(xlsm|xlsx)$", [Text.RegularExpressions.RegexOptions]::IgnoreCase)
  if (-not $NameMatch.Success) { throw "Excel filename must be newsletter_DB_YYYYMMDD.xlsm or .xlsx." }
  $IssueDate = [DateTime]::ParseExact($NameMatch.Groups[1].Value, "yyyyMMdd", [Globalization.CultureInfo]::InvariantCulture)
  if ($NewsletterDate) {
    $Explicit = [DateTime]::ParseExact($NewsletterDate, "yyyy-MM-dd", [Globalization.CultureInfo]::InvariantCulture)
    if ($Explicit.Date -ne $IssueDate.Date) { throw "Excel filename date and NewsletterDate do not match." }
  }
  [IO.Directory]::CreateDirectory($OutputDir) | Out-Null
  if (-not $AssetsDir) { $AssetsDir = Join-Path $OutputDir "images" }
  [IO.Directory]::CreateDirectory($AssetsDir) | Out-Null

  $Package = [IO.Compression.ZipFile]::OpenRead($XlsxPath)
  try {
    $Sheet = Read-TargetSheet $Package
    $Records = @(Get-Records $Sheet.Rows $IssueDate)
    $Images = @(Get-SheetImages $Package $Sheet.Part)
  } finally { $Package.Dispose() }

  $BannerRecords = @($Records | Where-Object { $_.section -eq "banner" } | Sort-Object { [int]$_.slot })
  for ($Index = 0; $Index -lt $BannerRecords.Count -and $Index -lt $Images.Count; $Index++) {
    if ($Images[$Index].Extension -ne ".png") {
      $Message = "Banner $($BannerRecords[$Index].slot) image must be PNG; found $($Images[$Index].Extension)."
      if ($BestEffort) { $Warnings.Add($Message); continue } else { throw $Message }
    }
    $ImageName = "banner_{0:00}.png" -f [int]$BannerRecords[$Index].slot
    $ImagePath = Join-Path $AssetsDir $ImageName
    [IO.File]::WriteAllBytes($ImagePath, $Images[$Index].Bytes)
    $Relative = $ImagePath.Substring(([IO.Path]::GetFullPath($OutputDir).TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar).Length).Replace("\", "/")
    $BannerRecords[$Index].image_path = $Relative
  }
  if ($Images.Count -lt $BannerRecords.Count) { $Warnings.Add("There are fewer embedded images than banner rows. Check image_path values in the CSV.") }

  $CsvPath = Join-Path $OutputDir "newsletter.csv"
  Write-NewsletterCsv $CsvPath $Records $IssueDate
  Write-Host "Worksheet: $($Sheet.Name)"
  Write-Host "CSV: $CsvPath"
  Write-Host ("Records: business {0}, academy {1}, banner {2}" -f @($Records | Where-Object section -eq "business").Count, @($Records | Where-Object section -eq "academy").Count, @($Records | Where-Object section -eq "banner").Count)
  Write-Host ("Banner images: {0}" -f [Math]::Min($Images.Count, $BannerRecords.Count))
  foreach ($Warning in $Warnings) { Write-Warning $Warning }
  if ($Warnings.Count -gt 0) { exit 2 }
  exit 0
} catch {
  Write-Error $_.Exception.Message
  exit 1
}
