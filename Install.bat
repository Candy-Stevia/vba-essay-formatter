@echo off
setlocal
title Academic Formatter Macro - Installer (Cai Dat)
echo ============================================================================
echo  ACADEMIC FORMATTER MACRO - AUTO INSTALLER (CAI DAT TU DONG)
echo ============================================================================
echo.

set "SRC_DOTM=%~dp0AcademicFormatter.dotm"
if not exist "%SRC_DOTM%" (
    echo [ERROR] Missing AcademicFormatter.dotm next to Install.bat.
    echo [LOI] Khong tim thay file AcademicFormatter.dotm cung thu muc.
    echo.
    pause
    exit /b 1
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$c = (Get-Content -LiteralPath '%~f0' -Raw) -split ('#__' + 'PS_START__'), 2; Invoke-Expression $c[1]"
set "EC=%errorlevel%"
echo.
if not "%EC%"=="0" (
    echo Installation stopped with code %EC%. ^(Cai dat dung voi ma loi %EC%.^)
)
pause
exit /b %EC%

#__PS_START__
$ErrorActionPreference = 'Stop'
$src = $env:SRC_DOTM
$wordExe = $null
$majorVer = $null

foreach ($r in @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe')) {
    if (Test-Path $r) {
        $p = (Get-ItemProperty $r -ErrorAction SilentlyContinue).'(default)'
        if ($p -and (Test-Path $p)) { $wordExe = $p; break }
    }
}

if (-not $wordExe) {
    foreach ($v in @('16.0','15.0','14.0','12.0')) {
        foreach ($base in @('HKLM:\SOFTWARE\Microsoft\Office', 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office')) {
            $k = "$base\$v\Word\InstallRoot"
            if (Test-Path $k) {
                $dir = (Get-ItemProperty $k -ErrorAction SilentlyContinue).Path
                if ($dir -and (Test-Path (Join-Path $dir 'WINWORD.EXE'))) {
                    $wordExe = Join-Path $dir 'WINWORD.EXE'
                    $majorVer = [int]($v.Split('.')[0])
                    break
                }
            }
        }
        if ($wordExe) { break }
    }
}

if (-not $wordExe) {
    Write-Host '[ERROR] Microsoft Word not found on this PC.' -ForegroundColor Red
    Write-Host '[LOI] Khong tim thay Microsoft Word tren may nay.' -ForegroundColor Red
    exit 2
}

if (-not $majorVer) {
    try { $majorVer = (Get-Item $wordExe).VersionInfo.FileMajorPart } catch { $majorVer = 16 }
}
Write-Host "Detected Microsoft Word v$majorVer ($wordExe)" -ForegroundColor Cyan

if ($majorVer -lt 15) {
    Write-Host '[ERROR] Word 2013 (v15) or newer is required. Your version is unsupported.' -ForegroundColor Red
    Write-Host '[LOI] Can Microsoft Word 2013 tro len. Phien ban hien tai khong ho tro.' -ForegroundColor Red
    exit 3
}

$verKey = "$majorVer.0"
if ($majorVer -ge 16) { $verKey = '16.0' }

$startupDir = $null
$optKey = "HKCU:\Software\Microsoft\Office\$verKey\Word\Options"
if (Test-Path $optKey) {
    $custom = (Get-ItemProperty $optKey -Name 'STARTUP-PATH' -ErrorAction SilentlyContinue).'STARTUP-PATH'
    if ($custom -and $custom.Trim().Length -gt 0) {
        $startupDir = [Environment]::ExpandEnvironmentVariables($custom.Trim())
    }
}
if (-not $startupDir) {
    $startupDir = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP'
}
if (-not (Test-Path $startupDir)) {
    New-Item -ItemType Directory -Path $startupDir -Force | Out-Null
}

$dest = Join-Path $startupDir 'AcademicFormatter.dotm'
$wordRunning = Get-Process -Name 'WINWORD' -ErrorAction SilentlyContinue

try {
    Copy-Item -Path $src -Destination $dest -Force
} catch {
    if ($wordRunning) {
        Write-Host '[ERROR] Word is open and locking AcademicFormatter.dotm. Close Word and run again.' -ForegroundColor Red
        Write-Host '[LOI] Word dang mo va khoa file. Hay dong Word roi chay lai Install.bat.' -ForegroundColor Red
        exit 4
    }
    Write-Host "[ERROR] Copy failed: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "[LOI] Khong the chep file: $($_.Exception.Message)" -ForegroundColor Red
    exit 5
}

try { Unblock-File -Path $dest -ErrorAction SilentlyContinue } catch {}

$srcHash = (Get-FileHash $src -Algorithm SHA256).Hash
$dstHash = (Get-FileHash $dest -Algorithm SHA256).Hash
if ($srcHash -ne $dstHash) {
    Write-Host '[ERROR] File verification failed (SHA-256 mismatch).' -ForegroundColor Red
    Write-Host '[LOI] Kiem tra toan ven file that bai.' -ForegroundColor Red
    exit 6
}

Write-Host "[OK] Installed to: $dest" -ForegroundColor Green
Write-Host "[OK] Da cai dat vao: $dest" -ForegroundColor Green

$vbaWarn = $null
foreach ($sk in @("HKCU:\Software\Policies\Microsoft\Office\$verKey\Word\Security", "HKCU:\Software\Microsoft\Office\$verKey\Word\Security")) {
    if (Test-Path $sk) {
        $w = (Get-ItemProperty $sk -Name 'VBAWarnings' -ErrorAction SilentlyContinue).VBAWarnings
        if ($null -ne $w) { $vbaWarn = [int]$w; break }
    }
}
if ($vbaWarn -eq 4) {
    Write-Host '[WARNING] Word Macro Settings are set to Disable all without notification.' -ForegroundColor Yellow
    Write-Host '          Open Word > File > Options > Trust Center > Macro Settings to enable.' -ForegroundColor Yellow
    Write-Host '[CANH BAO] Word dang tat tat ca Macro khong thong bao.' -ForegroundColor Yellow
    Write-Host '           Vao File > Options > Trust Center > Macro Settings de bat.' -ForegroundColor Yellow
}

if ($wordRunning) {
    Write-Host '[NOTE] Word is currently open. Restart Word to load the Formatter Macro tab.' -ForegroundColor Yellow
    Write-Host '[LUU Y] Word dang mo. Hay dong va mo lai Word de thay tab Formatter Macro.' -ForegroundColor Yellow
} else {
    Write-Host '[DONE] Open Word and select the Formatter Macro tab.' -ForegroundColor Green
    Write-Host '[XONG] Mo Word va chon tab Formatter Macro de su dung.' -ForegroundColor Green
}
exit 0
