[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)][string]$Csv,
  [Parameter(Mandatory = $true)][string]$Html,
  [Parameter(Mandatory = $true)][string]$Output,
  [switch]$BestEffort
)

$ErrorActionPreference = "Stop"
$UnsubscribeUrl = "https://www.kotra.or.kr/membership/mberEmailRecive.api?type="
$Prefixes = @{ business = "사업안내"; academy = "아카데미"; banner = "홍보배너" }
$Capacities = @{ business = 8; academy = 4; banner = 3 }
$Warnings = New-Object System.Collections.Generic.List[string]

function Get-AttrValue {
  param([string]$Tag, [string]$Name)
  $Match = [regex]::Match($Tag, "\b$([regex]::Escape($Name))\s*=\s*([`"'])(.*?)\1", "IgnoreCase,Singleline")
  if (-not $Match.Success) { return "" }
  return [Net.WebUtility]::HtmlDecode($Match.Groups[2].Value).Trim()
}

function Get-ManagedKey {
  param([string]$Alt)
  foreach ($Section in @("business", "academy", "banner")) {
    $Match = [regex]::Match($Alt, "^$([regex]::Escape($Prefixes[$Section]))(\d+)$")
    if ($Match.Success) {
      $Slot = [int]$Match.Groups[1].Value
      if ($Slot -ge 1 -and $Slot -le $Capacities[$Section]) {
        return "$Section|$Slot"
      }
    }
  }
  return $null
}

function Set-Href {
  param([string]$Tag, [string]$Url)
  $Pattern = "\s+href\s*=\s*([`"']).*?\1"
  $Regex = New-Object Text.RegularExpressions.Regex($Pattern, [Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [Text.RegularExpressions.RegexOptions]::Singleline)
  if ($Url) {
    $Escaped = [Net.WebUtility]::HtmlEncode($Url)
    $Attribute = ' href="' + $Escaped + '"'
    if ($Regex.IsMatch($Tag)) {
      return $Regex.Replace($Tag, [Text.RegularExpressions.MatchEvaluator]{ param($m) $Attribute }, 1)
    }
    return $Tag.Substring(0, $Tag.Length - 1) + $Attribute + ">"
  }
  return $Regex.Replace($Tag, "", 1)
}

function Read-CsvLinks {
  param([string]$Path)
  $Links = @{}
  $Rows = @(Import-Csv -LiteralPath $Path -Encoding UTF8)
  if ($Rows.Count -eq 0) { throw "CSV has no data rows." }
  $Names = @($Rows[0].PSObject.Properties.Name)
  foreach ($Required in @("section", "slot", "url")) {
    if ($Names -notcontains $Required) { throw "CSV requires section, slot, and url columns." }
  }
  $LineNo = 1
  foreach ($Row in $Rows) {
    $LineNo++
    $Section = ([string]$Row.section).Trim().ToLowerInvariant()
    if (-not $Prefixes.ContainsKey($Section)) { continue }
    [int]$Slot = 0
    if (-not [int]::TryParse(([string]$Row.slot).Trim(), [ref]$Slot)) {
      $Message = "CSV line $LineNo has a non-numeric slot and was skipped."
      if ($BestEffort) { $Warnings.Add($Message); continue } else { throw $Message }
    }
    if ($Slot -lt 1 -or $Slot -gt $Capacities[$Section]) {
      $Message = "CSV line $LineNo exceeds the $Section HTML capacity."
      if ($BestEffort) { $Warnings.Add($Message); continue } else { throw $Message }
    }
    $Url = (([string]$Row.url) -split "`r?`n", 2)[0].Trim()
    if ($Url -notmatch '^https?://') {
      $Message = "CSV line $LineNo has an invalid URL and was skipped: $Url"
      if ($BestEffort) { $Warnings.Add($Message); continue } else { throw $Message }
    }
    $Key = "$Section|$Slot"
    if ($Links.ContainsKey($Key)) {
      $Message = "CSV has a duplicate $Section slot $Slot; the later row was skipped."
      if ($BestEffort) { $Warnings.Add($Message); continue } else { throw $Message }
    }
    $Links[$Key] = $Url
  }
  return $Links
}

function Get-ProtectedHrefs {
  param([string]$Source)
  $Result = @{}
  $Index = 0
  foreach ($Match in [regex]::Matches($Source, "<area\b[^>]*>", "IgnoreCase,Singleline")) {
    $Index++
    $Tag = $Match.Value
    $Alt = Get-AttrValue $Tag "alt"
    if ((Get-ManagedKey $Alt) -or $Alt -eq "수신거부") { continue }
    $HrefMatch = [regex]::Match($Tag, "\s+href\s*=\s*([`"']).*?\1", "IgnoreCase,Singleline")
    $Result["$Index`:$Alt"] = if ($HrefMatch.Success) { $HrefMatch.Value } else { "" }
  }
  return $Result
}

function Test-HashtableEqual {
  param([hashtable]$Left, [hashtable]$Right)
  if ($Left.Count -ne $Right.Count) { return $false }
  foreach ($Key in $Left.Keys) {
    if (-not $Right.ContainsKey($Key) -or $Right[$Key] -ne $Left[$Key]) { return $false }
  }
  return $true
}

try {
  $Source = [IO.File]::ReadAllText((Resolve-Path -LiteralPath $Html).Path, [Text.Encoding]::UTF8)
  $ProtectedBefore = Get-ProtectedHrefs $Source
  try { $Links = Read-CsvLinks (Resolve-Path -LiteralPath $Csv).Path } catch {
    if (-not $BestEffort) { throw }
    $Links = @{}
    $Warnings.Add("Could not read CSV; existing links were kept: $($_.Exception.Message)")
  }

  $Counts = @{ business = 0; academy = 0; banner = 0; preserved = 0; unsubscribe = 0 }
  $AreaEvaluator = [Text.RegularExpressions.MatchEvaluator]{
    param($Match)
    $Tag = $Match.Value
    $Alt = Get-AttrValue $Tag "alt"
    if ($Alt -eq "수신거부") {
      $Counts.unsubscribe++
      return Set-Href $Tag $UnsubscribeUrl
    }
    $Key = Get-ManagedKey $Alt
    if (-not $Key) { return $Tag }
    if ($Key -eq "banner|2") { $Counts.preserved++; return $Tag }
    if ($Links.ContainsKey($Key)) {
      $Section = $Key.Split("|")[0]
      $Counts[$Section]++
      return Set-Href $Tag $Links[$Key]
    }
    $Counts.preserved++
    return $Tag
  }
  $AreaRegex = New-Object Text.RegularExpressions.Regex("<area\b[^>]*>", [Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [Text.RegularExpressions.RegexOptions]::Singleline)
  $Result = $AreaRegex.Replace($Source, $AreaEvaluator)
  if (-not $BestEffort -and $Counts.unsubscribe -ne 1) { throw 'Expected exactly one area with alt="수신거부".' }
  if ($BestEffort -and $Counts.unsubscribe -ne 1) { $Warnings.Add('Could not find exactly one area with alt="수신거부".') }

  $HtmlName = [IO.Path]::GetFileName($Html)
  $NameMatch = [regex]::Match($HtmlName, "^newsletter_(20\d{6})\.html$", "IgnoreCase")
  if (-not $NameMatch.Success) { throw "HTML filename must be newsletter_YYYYMMDD.html: $HtmlName" }
  $ImageName = "images/newsletter_$($NameMatch.Groups[1].Value).png"
  $ImageState = @{ Count = 0 }
  $ImageEvaluator = [Text.RegularExpressions.MatchEvaluator]{
    param($Match)
    $ImageState.Count++
    $Tag = $Match.Value
    $SrcPattern = "\bsrc\s*=\s*([`"']).*?\1"
    $Attr = 'src="' + $ImageName + '"'
    $SrcRegex = New-Object Text.RegularExpressions.Regex($SrcPattern, [Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [Text.RegularExpressions.RegexOptions]::Singleline)
    if ($SrcRegex.IsMatch($Tag)) {
      return $SrcRegex.Replace($Tag, [Text.RegularExpressions.MatchEvaluator]{ param($m) $Attr }, 1)
    }
    return $Tag.Substring(0, $Tag.Length - 1) + " $Attr>"
  }
  $MainImagePattern = "<img\b(?=[^>]*\busemap\s*=\s*([`"'])#Map\1)[^>]*>"
  $MainImageRegex = New-Object Text.RegularExpressions.Regex($MainImagePattern, [Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [Text.RegularExpressions.RegexOptions]::Singleline)
  $Result = $MainImageRegex.Replace($Result, $ImageEvaluator)
  if ($ImageState.Count -ne 1) {
    $Message = 'Expected exactly one main image using usemap="#Map".'
    if ($BestEffort) { $Warnings.Add($Message) } else { throw $Message }
  }
  if (-not (Test-HashtableEqual $ProtectedBefore (Get-ProtectedHrefs $Result))) {
    $Message = "A link outside the managed areas changed."
    if ($BestEffort) { $Warnings.Add($Message) } else { throw $Message }
  }

  $Parent = Split-Path -Parent $Output
  if ($Parent) { [IO.Directory]::CreateDirectory($Parent) | Out-Null }
  [IO.File]::WriteAllText($Output, $Result, (New-Object Text.UTF8Encoding($true)))
  Write-Host "VDI HTML: $Output"
  Write-Host ("Links: business {0}, academy {1}, banner {2}; preserved {3}; unsubscribe {4}" -f $Counts.business, $Counts.academy, $Counts.banner, $Counts.preserved, $Counts.unsubscribe)
  foreach ($Warning in $Warnings) { Write-Warning $Warning }
  if ($Warnings.Count -gt 0) { exit 2 }
  exit 0
} catch {
  Write-Error $_.Exception.Message
  exit 1
}
