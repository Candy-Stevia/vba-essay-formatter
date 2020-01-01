@echo off
setlocal
title Academic Formatter Macro - Install or Update (Cai Dat / Cap Nhat)
echo ============================================================================
echo  ACADEMIC FORMATTER MACRO - INSTALL OR UPDATE (CAI DAT / CAP NHAT TU DONG)
echo ============================================================================
echo.

set "SRC_DOTM=%~dp0AcademicFormatter.dotm"

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
$localSrc = $env:SRC_DOTM
$hasLocal = Test-Path $localSrc
$installSrc = $localSrc

$releaseUrl = 'https://github.com/Candy-Stevia/vba-essay-formatter/releases/latest/download/AcademicFormatter.dotm'
$rawMainUrl = 'https://raw.githubusercontent.com/Candy-Stevia/vba-essay-formatter/main/AcademicFormatter.dotm'
$manualUrl = 'https://github.com/Candy-Stevia/vba-essay-formatter/releases/latest'

$tmpFile = Join-Path $env:TEMP ('AF_update_' + [guid]::NewGuid().ToString('N') + '.dotm')
$downloadedOk = $false

Write-Host '[CHECK] Checking GitHub for latest version... (Dang kiem tra ban moi tren GitHub...)' -ForegroundColor Cyan

try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $urls = @($releaseUrl, $rawMainUrl)
    $lastEx = $null
    foreach ($u in $urls) {
        try {
            Invoke-WebRequest -Uri $u -OutFile $tmpFile -TimeoutSec 5 -UseBasicParsing
            $downloadedOk = $true
            break
        } catch {
            $lastEx = $_
            if ($_.Exception.Response -and $_.Exception.Response.StatusCode.value__ -eq 404) {
                continue
            }
            throw
        }
    }
    if (-not $downloadedOk -and $lastEx) { throw $lastEx }

    try { Unblock-File -Path $tmpFile -ErrorAction SilentlyContinue } catch {}
    if (-not (Test-Path $tmpFile)) {
        throw [System.UnauthorizedAccessException]::new('Downloaded file was removed or blocked by security software.')
    }
    $fi = Get-Item $tmpFile
    if ($fi.Length -lt 50000) {
        throw [System.InvalidOperationException]::new('Downloaded file is too small or invalid.')
    }
    $bytes = [System.IO.File]::ReadAllBytes($tmpFile)
    if ($bytes.Length -lt 2 -or $bytes[0] -ne 0x50 -or $bytes[1] -ne 0x4B) {
        throw [System.InvalidOperationException]::new('Downloaded file is not a valid .dotm archive.')
    }

    $remoteHash = (Get-FileHash $tmpFile -Algorithm SHA256).Hash
    $localHash = $null
    if ($hasLocal) {
        $localHash = (Get-FileHash $localSrc -Algorithm SHA256).Hash
    }

    if ($localHash -and $remoteHash -eq $localHash) {
        Write-Host '[UP TO DATE] You already have the latest version on this PC.' -ForegroundColor Green
        Write-Host '[MOI NHAT] Ban dang su dung phien ban moi nhat co san tren may.' -ForegroundColor Green
        $installSrc = $localSrc
    } else {
        Write-Host '[UPDATE] Newer version found on GitHub! Using downloaded AcademicFormatter.dotm.' -ForegroundColor Green
        Write-Host '[CAP NHAT] Da co phien ban moi hon tren GitHub! Dang nap ban moi.' -ForegroundColor Green
        $installSrc = $tmpFile
        if ($hasLocal) {
            try { Copy-Item -Path $tmpFile -Destination $localSrc -Force -ErrorAction SilentlyContinue } catch {}
        }
    }
} catch {
    $msg = $_.Exception.Message
    $isBlocked = ($_.Exception -is [System.UnauthorizedAccessException]) -or ($msg -match 'denied|blocked|virus|defender|antivirus|security|Unauthorized')
    if ($isBlocked) {
        Write-Host '[BLOCKED] Could not download new file (possibly blocked by Windows Defender or security policy).' -ForegroundColor Yellow
        Write-Host "          You can download the latest version manually at: $manualUrl" -ForegroundColor Yellow
        Write-Host '[BI CHAN] Khong tai duoc file moi (co the bi Windows Defender chan).' -ForegroundColor Yellow
        Write-Host "          Ban co the len $manualUrl de tai ban moi nhat." -ForegroundColor Yellow
    } else {
        Write-Host '[OFFLINE] Could not reach GitHub. Using local AcademicFormatter.dotm on this PC.' -ForegroundColor Yellow
        Write-Host '[NGOAI TUYEN] Khong ket noi duoc GitHub. Dang lay file co san tren may de nap vao Word.' -ForegroundColor Yellow
    }

    if (-not $hasLocal) {
        Write-Host '[ERROR] No local AcademicFormatter.dotm found next to install-or-update.bat.' -ForegroundColor Red
        Write-Host '[LOI] Khong tim thay file AcademicFormatter.dotm co san cung thu muc.' -ForegroundColor Red
        if (Test-Path $tmpFile) { Remove-Item $tmpFile -Force -ErrorAction SilentlyContinue }
        exit 1
    }
    $installSrc = $localSrc
}

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
    if (Test-Path $tmpFile) { Remove-Item $tmpFile -Force -ErrorAction SilentlyContinue }
    exit 2
}

if (-not $majorVer) {
    try { $majorVer = (Get-Item $wordExe).VersionInfo.FileMajorPart } catch { $majorVer = 16 }
}
Write-Host "Detected Microsoft Word v$majorVer ($wordExe)" -ForegroundColor Cyan

if ($majorVer -lt 15) {
    Write-Host '[ERROR] Word 2013 (v15) or newer is required. Your version is unsupported.' -ForegroundColor Red
    Write-Host '[LOI] Can Microsoft Word 2013 tro len. Phien ban hien tai khong ho tro.' -ForegroundColor Red
    if (Test-Path $tmpFile) { Remove-Item $tmpFile -Force -ErrorAction SilentlyContinue }
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
    Copy-Item -Path $installSrc -Destination $dest -Force
} catch {
    if ($tmpFile -and (Test-Path $tmpFile)) { Remove-Item $tmpFile -Force -ErrorAction SilentlyContinue }
    if ($wordRunning) {
        Write-Host '[ERROR] Word is open and locking AcademicFormatter.dotm. Close Word and run again.' -ForegroundColor Red
        Write-Host '[LOI] Word dang mo va khoa file. Hay dong Word roi chay lai install-or-update.bat.' -ForegroundColor Red
        exit 4
    }
    Write-Host "[ERROR] Copy failed: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "[LOI] Khong the chep file: $($_.Exception.Message)" -ForegroundColor Red
    exit 5
}

try { Unblock-File -Path $dest -ErrorAction SilentlyContinue } catch {}

$srcHash = (Get-FileHash $installSrc -Algorithm SHA256).Hash
$dstHash = (Get-FileHash $dest -Algorithm SHA256).Hash
if (Test-Path $tmpFile) { Remove-Item $tmpFile -Force -ErrorAction SilentlyContinue }

if ($srcHash -ne $dstHash) {
    Write-Host '[ERROR] File verification failed (SHA-256 mismatch).' -ForegroundColor Red
    Write-Host '[LOI] Kiem tra toan ven file that bai.' -ForegroundColor Red
    exit 6
}

Write-Host "[OK] Installed to: $dest" -ForegroundColor Green
Write-Host "[OK] Da cai dat vao: $dest" -ForegroundColor Green

try {
    foreach ($ov in @('16.0','15.0')) {
        $disKey = "HKCU:\Software\Microsoft\Office\$ov\Word\Resiliency\DisabledItems"
        if (Test-Path $disKey) {
            $props = Get-ItemProperty $disKey -ErrorAction SilentlyContinue
            foreach ($pr in $props.PSObject.Properties) {
                if ($pr.Name -notlike 'PS*') {
                    $txt = [System.Text.Encoding]::Unicode.GetString([byte[]]$pr.Value) -replace "`0", ''
                    if ($txt -match 'AcademicFormatter\.dotm') { Remove-ItemProperty -Path $disKey -Name $pr.Name -ErrorAction SilentlyContinue }
                }
            }
        }
    }
} catch {}

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