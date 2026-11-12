[CmdletBinding()]
param([Parameter(Mandatory = $true)][string]$OutputDir)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

try {
  $ResolvedOutput = (Resolve-Path -LiteralPath $OutputDir).Path
  if (-not (Test-Path -LiteralPath $ResolvedOutput -PathType Container)) {
    throw "Output folder not found: $OutputDir"
  }
  $Issue8 = Split-Path (Split-Path $ResolvedOutput -Parent) -Leaf
  if ($Issue8 -notmatch '^20\d{6}$') { throw "The output parent folder must be YYYYMMDD: $Issue8" }
  $null = [DateTime]::ParseExact($Issue8, "yyyyMMdd", [Globalization.CultureInfo]::InvariantCulture)

  $ImageNames = @(
    "bg.png", "banner_01.png", "banner_02.png", "banner_03.png",
    "우측배너.png", "newsletter_${Issue8}.png"
  )
  $ImagesDir = Join-Path $ResolvedOutput "images"
  $VdiHtml = Join-Path $ResolvedOutput "vdi_${Issue8}.html"
  $MutuHtml = Join-Path $ResolvedOutput "mutu24_${Issue8}.html"
  $ZipPath = Join-Path $ResolvedOutput "${Issue8}_반출.zip"
  $Required = @($ImageNames | ForEach-Object { Join-Path $ImagesDir $_ }) + @($VdiHtml, $MutuHtml)
  $Missing = @($Required | Where-Object { -not (Test-Path -LiteralPath $_ -PathType Leaf) })
  if ($Missing.Count -gt 0) { throw "Required ZIP files are missing: $($Missing -join ', ')" }

  if (Test-Path -LiteralPath $ZipPath) { Remove-Item -LiteralPath $ZipPath -Force }
  $Archive = [IO.Compression.ZipFile]::Open($ZipPath, [IO.Compression.ZipArchiveMode]::Create)
  try {
    foreach ($ImageName in $ImageNames) {
      [IO.Compression.ZipFileExtensions]::CreateEntryFromFile($Archive, (Join-Path $ImagesDir $ImageName), "images/$ImageName", [IO.Compression.CompressionLevel]::Optimal) | Out-Null
    }
    [IO.Compression.ZipFileExtensions]::CreateEntryFromFile($Archive, $VdiHtml, [IO.Path]::GetFileName($VdiHtml), [IO.Compression.CompressionLevel]::Optimal) | Out-Null
    [IO.Compression.ZipFileExtensions]::CreateEntryFromFile($Archive, $MutuHtml, [IO.Path]::GetFileName($MutuHtml), [IO.Compression.CompressionLevel]::Optimal) | Out-Null
  } finally { $Archive.Dispose() }
  Write-Host "Transfer ZIP: $ZipPath"
  exit 0
} catch {
  Write-Error $_.Exception.Message
  exit 1
}
