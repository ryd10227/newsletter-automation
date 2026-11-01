$WorkDir = (Get-Location).Path
$Root = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$Issue8 = Split-Path $WorkDir -Leaf

if ($Issue8 -notmatch '^20\d{6}$') {
  throw "Open PowerShell inside work_history\YYYYMMDD. Current folder: $Issue8"
}

try {
  $IssueDate = [DateTime]::ParseExact(
    $Issue8,
    "yyyyMMdd",
    [Globalization.CultureInfo]::InvariantCulture
  ).ToString("yyyy-MM-dd")
} catch {
  throw "The work folder date is invalid: $Issue8"
}

$ExpectedWorkRoot = (Resolve-Path (Join-Path $Root "work_history")).Path
if ((Split-Path $WorkDir -Parent) -ne $ExpectedWorkRoot) {
  throw "The date folder must be directly under work_history."
}

$InputDir = Join-Path $WorkDir "input"
$OutputDir = Join-Path $WorkDir "output"
$ImagesDir = Join-Path $OutputDir "images"

if (-not (Test-Path -LiteralPath $InputDir -PathType Container)) {
  throw "Missing input folder: $InputDir"
}

$ExcelFiles = @(
  Get-ChildItem -LiteralPath $InputDir -File |
    Where-Object { $_.Extension -ieq ".xlsm" -or $_.Extension -ieq ".xlsx" }
)
$ExpectedExcelPattern = "^newsletter_DB_${Issue8}\.(xlsm|xlsx)$"
if ($ExcelFiles.Count -ne 1 -or $ExcelFiles[0].Name -notmatch $ExpectedExcelPattern) {
  throw "Keep exactly one Excel file named newsletter_DB_${Issue8}.xlsm or .xlsx in input."
}
$ExcelFile = $ExcelFiles[0]

$NewsletterDirs = @(
  Get-ChildItem -LiteralPath $InputDir -Directory |
    Where-Object { $_.Name -match '^20\d{6}_letter$' }
)
$ExpectedNewsletterName = "${Issue8}_letter"
if ($NewsletterDirs.Count -ne 1 -or $NewsletterDirs[0].Name -ne $ExpectedNewsletterName) {
  throw "Keep exactly one source folder named $ExpectedNewsletterName in input."
}
$NewsletterDir = $NewsletterDirs[0].FullName

$Psd = Join-Path $NewsletterDir "newsletter_${Issue8}.psd"
$SourceHtml = Join-Path $NewsletterDir "newsletter_${Issue8}.html"
foreach ($RequiredFile in @($Psd, $SourceHtml)) {
  if (-not (Test-Path -LiteralPath $RequiredFile -PathType Leaf)) {
    throw "Missing required file: $RequiredFile"
  }
}

$OtherCoreFiles = @(
  Get-ChildItem -LiteralPath $NewsletterDir -File |
    Where-Object {
      ($_.Extension -ieq ".psd" -or $_.Extension -ieq ".html") -and
      $_.FullName -ne $Psd -and $_.FullName -ne $SourceHtml
    }
)
if ($OtherCoreFiles.Count -gt 0) {
  throw "The source folder contains an extra PSD or HTML file."
}

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
New-Item -ItemType Directory -Force -Path $ImagesDir | Out-Null

$Excel = $ExcelFile.FullName
$Csv = Join-Path $OutputDir "newsletter.csv"
$ResultPsd = Join-Path $OutputDir "newsletter_${Issue8}.psd"
$VdiOutput = Join-Path $OutputDir "vdi_${Issue8}.html"
$MutuOutput = Join-Path $OutputDir "mutu24_${Issue8}.html"
$MutuTemplate = Join-Path $Root "3_html\mutu24_template.txt"
$PhotoshopScript = Join-Path $Root "2_photoshop\fill.jsx"
$PhotoshopRunner = Join-Path $OutputDir "_run_photoshop_newsletter.jsx"
$PhotoshopLog = Join-Path $OutputDir "photoshop_automation_error.log"
$WebPng = Join-Path $ImagesDir "newsletter_${Issue8}.png"

function ConvertTo-JsPath {
  param([string]$Value)
  return $Value.Replace('\', '/').Replace('"', '\"')
}

function Invoke-PhotoshopAutomation {
  if (-not (Test-Path -LiteralPath $Csv -PathType Leaf)) {
    Write-Host "Missing CSV. Enter 1 before entering 2." -ForegroundColor Yellow
    return
  }

  if (-not (Test-Path -LiteralPath $PhotoshopScript -PathType Leaf)) {
    Write-Host "Missing Photoshop script: $PhotoshopScript" -ForegroundColor Yellow
    return
  }

  $RunnerLines = @(
    "#target photoshop",
    ('var NEWSLETTER_TEMPLATE_PATH = "{0}";' -f (ConvertTo-JsPath $Psd)),
    ('$' + '.evalFile(new File("{0}"));' -f (ConvertTo-JsPath $PhotoshopScript))
  )

  Remove-Item -LiteralPath $PhotoshopLog -Force -ErrorAction SilentlyContinue
  Set-Content -LiteralPath $PhotoshopRunner -Value $RunnerLines -Encoding Default

  try {
    Write-Host "Starting Photoshop automation. Please wait..."
    $PhotoshopApplication = New-Object -ComObject "Photoshop.Application"
    $PhotoshopApplication.Visible = $true
    $PhotoshopApplication.DoJavaScriptFile($PhotoshopRunner)

    if (Test-Path -LiteralPath $PhotoshopLog -PathType Leaf) {
      $PhotoshopStatus = Get-Content -LiteralPath $PhotoshopLog -TotalCount 1
      if ($PhotoshopStatus -eq "SUCCESS") {
        Remove-Item -LiteralPath $PhotoshopLog -Force -ErrorAction SilentlyContinue
        Write-Host "Created: $ResultPsd" -ForegroundColor Green
        Write-Host "Created: $WebPng" -ForegroundColor Green
      } else {
        Write-Host "Photoshop finished with warnings or errors: $PhotoshopLog" -ForegroundColor Yellow
      }
    } else {
      Write-Host "Photoshop finished, but no status log was returned. Check the output folder." -ForegroundColor Yellow
    }
  } catch {
    $PhotoshopError = "FAILED`r`n" + $_.Exception.Message
    Set-Content -LiteralPath $PhotoshopLog -Value $PhotoshopError -Encoding UTF8
    Write-Host "Photoshop automation failed: $PhotoshopLog" -ForegroundColor Yellow

    $PhotoshopExecutable = Find-PhotoshopExecutable
    if ($PhotoshopExecutable) {
      Start-Process -FilePath $PhotoshopExecutable -ArgumentList ('"{0}"' -f $Psd)
      Write-Host "The source PSD was opened for manual recovery." -ForegroundColor Yellow
    }
  } finally {
    Remove-Item -LiteralPath $PhotoshopRunner -Force -ErrorAction SilentlyContinue
  }
}

function Find-PhotoshopExecutable {
  $RegistryPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Photoshop.exe",
    "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Photoshop.exe",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\Photoshop.exe"
  )

  foreach ($RegistryPath in $RegistryPaths) {
    $RegistryEntry = Get-ItemProperty -LiteralPath $RegistryPath -ErrorAction SilentlyContinue
    if ($null -ne $RegistryEntry) {
      $RegistryExecutable = $RegistryEntry.'(default)'
      if ($RegistryExecutable -and (Test-Path -LiteralPath $RegistryExecutable -PathType Leaf)) {
        return $RegistryExecutable
      }
    }
  }

  $PhotoshopCommand = Get-Command "Photoshop.exe" -ErrorAction SilentlyContinue
  if ($null -ne $PhotoshopCommand -and $PhotoshopCommand.Source) {
    return $PhotoshopCommand.Source
  }

  $AdobeRoot = Join-Path $env:ProgramFiles "Adobe"
  if (Test-Path -LiteralPath $AdobeRoot -PathType Container) {
    $PhotoshopFolders = @(
      Get-ChildItem -LiteralPath $AdobeRoot -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -like "Adobe Photoshop*" } |
        Sort-Object Name -Descending
    )
    foreach ($PhotoshopFolder in $PhotoshopFolders) {
      $PhotoshopExecutable = Join-Path $PhotoshopFolder.FullName "Photoshop.exe"
      if (Test-Path -LiteralPath $PhotoshopExecutable -PathType Leaf) {
        return $PhotoshopExecutable
      }
    }
  }

  return $null
}

Write-Host "Ready: $Issue8"
Write-Host "Input:  $InputDir"
Write-Host "Output: $OutputDir"

function Show-NewsletterMenu {
  Write-Host ""
  Write-Host "========================================"
  Write-Host "  Newsletter automation: $Issue8"
  Write-Host "========================================"
  Write-Host "  1  Excel -> CSV"
  Write-Host "  2  Run Photoshop automation"
  Write-Host "  3  VDI HTML + Mutu24 HTML"
  Write-Host "  4  Create transfer ZIP"
  Write-Host "  0  Exit"
  Write-Host "========================================"
}

while ($true) {
  Show-NewsletterMenu
  $Choice = Read-Host "Enter a number"

  switch ($Choice.Trim()) {
    "1" {
      & py "$Root\1_excel_to_csv\export.py" `
        --xlsx "$Excel" `
        --output-dir "$OutputDir" `
        --assets-dir "$ImagesDir" `
        --best-effort

      if ($LASTEXITCODE -eq 0) {
        Write-Host "Created: $Csv" -ForegroundColor Green
      } else {
        Write-Host "Step 1 finished with an error. Check the messages above." -ForegroundColor Yellow
      }
    }

    "2" {
      Write-Host ""
      Invoke-PhotoshopAutomation
      if ((Test-Path -LiteralPath $ResultPsd -PathType Leaf) -and (Test-Path -LiteralPath $WebPng -PathType Leaf)) {
        Write-Host "============================================================" -ForegroundColor Cyan
        Write-Host "[검수 필수] 2단계 Photoshop 결과" -ForegroundColor Cyan
        Write-Host "============================================================" -ForegroundColor Cyan
        Write-Host "[ ] 배너 1·3의 모서리가 둥근지 확인하세요."
        Write-Host "[ ] images\bg.png가 newsletter_${Issue8}.psd의 해외시장뉴스 영역의 배너와 일치하는지 확인하세요. 컬러배경·오브제·텍스트가 모두 포함된 이미지여야 합니다."
      }
    }

    "3" {
      & py "$Root\3_html\vdi.py" `
        --csv "$Csv" `
        --html "$SourceHtml" `
        --output "$VdiOutput" `
        --best-effort
      $VdiExitCode = $LASTEXITCODE

      & py "$Root\3_html\mutu24.py" `
        --template "$MutuTemplate" `
        --csv "$Csv" `
        --psd "$ResultPsd" `
        --map-html "$VdiOutput" `
        --output "$MutuOutput" `
        --issue-date "$IssueDate" `
        --best-effort
      $MutuExitCode = $LASTEXITCODE

      if ($VdiExitCode -eq 0 -and $MutuExitCode -eq 0) {
        Write-Host "Created: $VdiOutput" -ForegroundColor Green
        Write-Host "Created: $MutuOutput" -ForegroundColor Green
        Write-Host "============================================================" -ForegroundColor Cyan
        Write-Host "[검수 필수] 3단계 HTML 결과" -ForegroundColor Cyan
        Write-Host "============================================================" -ForegroundColor Cyan
        Write-Host "[ ] 두 HTML을 각각 열어 모든 요소가 올바른 URL로 연결되는지 확인하세요."
        Write-Host "[ ] 모든 URL을 직접 열고, 열린 페이지 안에 네이버폼·구글폼·모아폼 유도 링크나 QR코드가 없는지 확인하세요. 만약 있다면, 해당 글을 이번 뉴스레터에서 제거해야 합니다."
      } else {
        Write-Host "Step 3 finished with an error. Check the messages above." -ForegroundColor Yellow
      }
    }

    "4" {
      & py "$Root\4_create_zip\create.py" --output-dir "$OutputDir"

      if ($LASTEXITCODE -ne 0) {
        Write-Host "Step 4 finished with an error. Check the messages above." -ForegroundColor Yellow
      }
    }

    "0" {
      Write-Host "Finished."
      return
    }

    default {
      Write-Host "Enter 1, 2, 3, 4, or 0." -ForegroundColor Yellow
    }
  }
}
