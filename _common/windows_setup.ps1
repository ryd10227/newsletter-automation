$WorkDir = (Get-Location).Path
$Root = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$Issue8 = Split-Path $WorkDir -Leaf

if ($Issue8 -notmatch '^20\d{6}$') {
  throw "work_history\YYYYMMDD 폴더에서 실행해야 합니다. 현재 폴더: $Issue8"
}

try {
  $IssueDate = [DateTime]::ParseExact(
    $Issue8,
    "yyyyMMdd",
    [Globalization.CultureInfo]::InvariantCulture
  ).ToString("yyyy-MM-dd")
} catch {
  throw "작업 폴더 날짜가 올바르지 않습니다: $Issue8"
}

$ExpectedWorkRoot = (Resolve-Path (Join-Path $Root "work_history")).Path
if ((Split-Path $WorkDir -Parent) -ne $ExpectedWorkRoot) {
  throw "날짜 폴더는 work_history 바로 아래에 있어야 합니다."
}

$InputDir = Join-Path $WorkDir "input"
$OutputDir = Join-Path $WorkDir "output"
$ImagesDir = Join-Path $OutputDir "images"

if (-not (Test-Path -LiteralPath $InputDir -PathType Container)) {
  throw "input 폴더가 없습니다: $InputDir"
}

$ExcelFiles = @(
  Get-ChildItem -LiteralPath $InputDir -File |
    Where-Object { $_.Extension -ieq ".xlsm" -or $_.Extension -ieq ".xlsx" }
)
$ExpectedExcelPattern = "^newsletter_DB_${Issue8}\.(xlsm|xlsx)$"
if ($ExcelFiles.Count -ne 1 -or $ExcelFiles[0].Name -notmatch $ExpectedExcelPattern) {
  throw "input에는 newsletter_DB_${Issue8}.xlsm 또는 .xlsx 이름의 Excel 파일을 정확히 하나만 넣으세요."
}
$ExcelFile = $ExcelFiles[0]

$NewsletterDirs = @(
  Get-ChildItem -LiteralPath $InputDir -Directory |
    Where-Object { $_.Name -match '^20\d{6}_letter$' }
)
$ExpectedNewsletterName = "${Issue8}_letter"
if ($NewsletterDirs.Count -ne 1 -or $NewsletterDirs[0].Name -ne $ExpectedNewsletterName) {
  throw "input에는 $ExpectedNewsletterName 원본 폴더를 정확히 하나만 넣으세요."
}
$NewsletterDir = $NewsletterDirs[0].FullName

$Psd = Join-Path $NewsletterDir "newsletter_${Issue8}.psd"
$SourceHtml = Join-Path $NewsletterDir "newsletter_${Issue8}.html"
foreach ($RequiredFile in @($Psd, $SourceHtml)) {
  if (-not (Test-Path -LiteralPath $RequiredFile -PathType Leaf)) {
    throw "필수 파일이 없습니다: $RequiredFile"
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
  throw "원본 폴더에 날짜와 이름이 다른 PSD 또는 HTML 파일이 추가로 들어 있습니다."
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
    Write-Host "newsletter.csv가 없습니다. 먼저 1을 실행하세요." -ForegroundColor Yellow
    return 1
  }

  if (-not (Test-Path -LiteralPath $PhotoshopScript -PathType Leaf)) {
    Write-Host "Photoshop 스크립트가 없습니다: $PhotoshopScript" -ForegroundColor Yellow
    return 1
  }

  $RunnerLines = @(
    "#target photoshop",
    ('var NEWSLETTER_TEMPLATE_PATH = "{0}";' -f (ConvertTo-JsPath $Psd)),
    ('$' + '.evalFile(new File("{0}"));' -f (ConvertTo-JsPath $PhotoshopScript))
  )

  Remove-Item -LiteralPath $PhotoshopLog -Force -ErrorAction SilentlyContinue
  Set-Content -LiteralPath $PhotoshopRunner -Value $RunnerLines -Encoding Default

  try {
    Write-Host "Photoshop 자동화를 시작합니다. 완료될 때까지 기다리세요."
    $PhotoshopApplication = New-Object -ComObject "Photoshop.Application"
    $PhotoshopApplication.Visible = $true
    $null = $PhotoshopApplication.DoJavaScriptFile($PhotoshopRunner)

    if (Test-Path -LiteralPath $PhotoshopLog -PathType Leaf) {
      $PhotoshopStatus = Get-Content -LiteralPath $PhotoshopLog -TotalCount 1
      if ($PhotoshopStatus -eq "SUCCESS") {
        Remove-Item -LiteralPath $PhotoshopLog -Force -ErrorAction SilentlyContinue
        Write-Host "생성 완료: $ResultPsd" -ForegroundColor Green
        Write-Host "생성 완료: $WebPng" -ForegroundColor Green
        return 0
      } else {
        Write-Host "Photoshop 작업에 경고 또는 오류가 있습니다: $PhotoshopLog" -ForegroundColor Yellow
        return 2
      }
    } else {
      Write-Host "Photoshop 작업은 끝났지만 상태 로그를 확인할 수 없습니다. output 폴더를 확인하세요." -ForegroundColor Yellow
      return 2
    }
  } catch {
    $PhotoshopError = "FAILED`r`n" + $_.Exception.Message
    Set-Content -LiteralPath $PhotoshopLog -Value $PhotoshopError -Encoding UTF8
    Write-Host "Photoshop 자동화에 실패했습니다: $PhotoshopLog" -ForegroundColor Yellow

    $PhotoshopExecutable = Find-PhotoshopExecutable
    if ($PhotoshopExecutable) {
      Start-Process -FilePath $PhotoshopExecutable -ArgumentList ('"{0}"' -f $Psd)
      Write-Host "수동 복구를 위해 원본 PSD를 열었습니다." -ForegroundColor Yellow
    }
    return 1
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

function Wait-ForVerification {
  param(
    [string]$StepTitle,
    [string[]]$CheckItems
  )

  Write-Host ""
  Write-Host "============================================================" -ForegroundColor Cyan
  Write-Host "[검수 필수] $StepTitle" -ForegroundColor Cyan
  Write-Host "============================================================" -ForegroundColor Cyan
  foreach ($CheckItem in $CheckItems) {
    Write-Host "[ ] $CheckItem"
  }
  Write-Host ""
  $null = Read-Host "위 항목을 모두 확인한 뒤 Enter를 누르세요"
}

function Show-Step1Verification {
  Wait-ForVerification "1단계 Excel → CSV" @(
    "output\newsletter.csv가 생성되었는지 확인",
    "CSV의 사업안내·아카데미·배너 제목, 기간, 마감일, URL이 Excel과 일치하는지 확인",
    "output\images\banner_01.png와 banner_03.png가 올바른 이미지인지 확인",
    "경고가 표시됐다면 해당 행과 image_path를 반드시 확인"
  )
}

function Show-Step2Verification {
  Wait-ForVerification "2단계 Photoshop 결과" @(
    "output\newsletter_${Issue8}.psd와 output\images\newsletter_${Issue8}.png가 저장되었는지 확인",
    "결과 PSD와 전체 PNG를 열어 배너 1·3의 모서리가 둥글게 잘렸는지 확인",
    "배너 1·3이 기존 이미지에 가려지지 않고 올바른 이미지로 교체됐는지 확인",
    "사업안내·아카데미의 제목, 기간, 신청마감 텍스트가 잘리지 않고 정확한지 확인",
    "output\images\bg.png가 해외시장뉴스 영역 상단 배너 전체인지 확인",
    "bg.png에 PSD에서 보이는 배경·오브제·텍스트가 전부 포함됐는지 확인",
    "banner_02.png, 우측배너.png와 전체 뉴스레터 PNG가 정상인지 확인",
    "photoshop_automation_error.log가 있다면 내용을 확인하고 오류를 해결"
  )
}

function Show-Step3Verification {
  Wait-ForVerification "3단계 VDI·무투24 HTML" @(
    "output\vdi_${Issue8}.html과 output\mutu24_${Issue8}.html을 각각 열어 확인",
    "해외시장뉴스·사업안내·아카데미·배너·경제외교 등 모든 영역의 텍스트와 링크가 일치하는지 확인",
    "링크 순서가 화면의 항목 순서와 정확히 일치하는지 확인",
    "모든 URL을 하나씩 직접 열어 제목과 실제 연결 페이지 내용이 일치하는지 확인",
    "KOTRA 주소라도 페이지 안의 신청 버튼·연결 링크를 눌러 최종 목적지를 확인",
    "본문과 연결 페이지에 네이버폼·구글폼·모아폼 이용을 유도하는 링크가 없는지 확인",
    "본문과 이미지 안의 QR코드가 외부 설문조사 플랫폼으로 연결되지 않는지 확인",
    "금지 플랫폼이 발견되면 사용을 중단하고 지원요청자에게 KOTRA 내부 플랫폼 URL로 수정을 요청"
  )
}

function Show-Step4Verification {
  Wait-ForVerification "4단계 반입 ZIP" @(
    "output\${Issue8}_반입.zip이 생성되었는지 확인",
    "ZIP 안에 images 폴더, vdi_${Issue8}.html, mutu24_${Issue8}.html이 들어 있는지 확인",
    "최종 검수가 끝난 파일만 외부메일로 반입"
  )
}

Write-Host "준비 완료: $Issue8"
Write-Host "입력 폴더: $InputDir"
Write-Host "출력 폴더: $OutputDir"

function Show-NewsletterMenu {
  Write-Host ""
  Write-Host "========================================"
  Write-Host "  뉴스레터 자동화: $Issue8"
  Write-Host "========================================"
  Write-Host "  1  Excel → CSV"
  Write-Host "  2  Photoshop 전체 자동화"
  Write-Host "  3  VDI HTML + 무투24 HTML"
  Write-Host "  4  반입 ZIP 생성"
  Write-Host "  0  종료"
  Write-Host "========================================"
}

while ($true) {
  Show-NewsletterMenu
  $Choice = Read-Host "번호 입력"

  switch ($Choice.Trim()) {
    "1" {
      & py "$Root\1_excel_to_csv\export.py" `
        --xlsx "$Excel" `
        --output-dir "$OutputDir" `
        --assets-dir "$ImagesDir" `
        --best-effort

      if ($LASTEXITCODE -eq 0) {
        Write-Host "생성 완료: $Csv" -ForegroundColor Green
        Show-Step1Verification
      } elseif ($LASTEXITCODE -eq 2) {
        Write-Host "1단계가 경고와 함께 완료됐습니다. 위 경고를 확인하세요." -ForegroundColor Yellow
        Show-Step1Verification
      } else {
        Write-Host "1단계에서 오류가 발생했습니다. 위 메시지를 확인하세요." -ForegroundColor Yellow
      }
    }

    "2" {
      Write-Host ""
      $PhotoshopExitCode = Invoke-PhotoshopAutomation
      if ($PhotoshopExitCode -eq 0 -or $PhotoshopExitCode -eq 2) {
        Show-Step2Verification
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
        Write-Host "생성 완료: $VdiOutput" -ForegroundColor Green
        Write-Host "생성 완료: $MutuOutput" -ForegroundColor Green
        Show-Step3Verification
      } elseif (($VdiExitCode -eq 0 -or $VdiExitCode -eq 2) -and ($MutuExitCode -eq 0 -or $MutuExitCode -eq 2)) {
        Write-Host "3단계가 경고와 함께 완료됐습니다. 위 경고를 확인하세요." -ForegroundColor Yellow
        Show-Step3Verification
      } else {
        Write-Host "3단계에서 오류가 발생했습니다. 위 메시지를 확인하세요." -ForegroundColor Yellow
      }
    }

    "4" {
      & py "$Root\4_create_zip\create.py" --output-dir "$OutputDir"

      if ($LASTEXITCODE -eq 0) {
        Show-Step4Verification
      } else {
        Write-Host "4단계에서 오류가 발생했습니다. 위 메시지를 확인하세요." -ForegroundColor Yellow
      }
    }

    "0" {
      Write-Host "작업을 종료합니다."
      return
    }

    default {
      Write-Host "1, 2, 3, 4 또는 0을 입력하세요." -ForegroundColor Yellow
    }
  }
}
