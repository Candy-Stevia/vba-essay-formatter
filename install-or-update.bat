@echo off
REM [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
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
# [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Test-DotmIntegrity {
    param([string]$FilePath)
    if (-not (Test-Path -LiteralPath $FilePath)) { return $false }
    $fi = Get-Item -LiteralPath $FilePath -ErrorAction SilentlyContinue
    if (-not $fi -or $fi.Length -lt 50000) { return $false }
    $bytes = [System.IO.File]::ReadAllBytes($FilePath)
    if ($bytes.Length -lt 2 -or $bytes[0] -ne 0x50 -or $bytes[1] -ne 0x4B) { return $false }
    $zip = $null
    try {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($FilePath)
        $coreEntry = $zip.GetEntry('docProps/core.xml')
        $uiEntry = $zip.GetEntry('customUI/customUI14.xml')
        if (-not $coreEntry -or -not $uiEntry) { return $false }
        foreach ($entry in @($coreEntry, $uiEntry)) {
            $sr = New-Object System.IO.StreamReader($entry.Open())
            try { $xmlText = $sr.ReadToEnd() } finally { $sr.Dispose() }
            $docXml = New-Object System.Xml.XmlDocument
            $docXml.LoadXml($xmlText)
        }
        return $true
    } catch {
        return $false
    } finally {
        if ($zip) { $zip.Dispose() }
    }
}

function Clear-WordResiliencyBlock {
    $cleared = 0
    foreach ($ov in @('16.0', '15.0', '14.0', '12.0')) {
        foreach ($sub in @('DisabledItems', 'StartupItems')) {
            $regKey = "HKCU:\Software\Microsoft\Office\$ov\Word\Resiliency\$sub"
            if (Test-Path $regKey) {
                $props = Get-ItemProperty $regKey -ErrorAction SilentlyContinue
                foreach ($pr in $props.PSObject.Properties) {
                    if ($pr.Name -notlike 'PS*') {
                        try {
                            $txt = [System.Text.Encoding]::Unicode.GetString([byte[]]$pr.Value) -replace "`0", ''
                            if ($txt -match 'AcademicFormatter\.dotm') {
                                Remove-ItemProperty -Path $regKey -Name $pr.Name -ErrorAction SilentlyContinue
                                $cleared++
                            }
                        } catch {}
                    }
                }
            }
        }
    }
    return $cleared
}

$localSrc = $env:SRC_DOTM
$hasLocal = Test-Path -LiteralPath $localSrc
$localValid = $hasLocal -and (Test-DotmIntegrity -FilePath $localSrc)
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
    if (-not (Test-Path -LiteralPath $tmpFile)) {
        throw [System.UnauthorizedAccessException]::new('Downloaded file was removed or blocked by security software.')
    }
    if (-not (Test-DotmIntegrity -FilePath $tmpFile)) {
        throw [System.InvalidOperationException]::new('CORRUPT_XML: Downloaded .dotm failed ZIP/XML integrity verification.')
    }

    $remoteHash = (Get-FileHash $tmpFile -Algorithm SHA256).Hash
    $localHash = $null
    if ($localValid) {
        $localHash = (Get-FileHash $localSrc -Algorithm SHA256).Hash
    }

    if ($localHash -and $remoteHash -eq $localHash) {
        Write-Host '[UP TO DATE] You already have the latest version on this PC.' -ForegroundColor Green
        Write-Host '[MOI NHAT] Ban dang su dung phien ban moi nhat co san tren may.' -ForegroundColor Green
        $installSrc = $localSrc
    } else {
        Write-Host '[UPDATE] Newer valid version found on GitHub! Using downloaded AcademicFormatter.dotm.' -ForegroundColor Green
        Write-Host '[CAP NHAT] Da co phien ban moi hop le tren GitHub! Dang nap ban moi.' -ForegroundColor Green
        $installSrc = $tmpFile
        if ($hasLocal) {
            try { Copy-Item -Path $tmpFile -Destination $localSrc -Force -ErrorAction SilentlyContinue } catch {}
        }
    }
} catch {
    $msg = $_.Exception.Message
    $isCorrupt = $msg -match 'CORRUPT_XML'
    $isBlocked = ($_.Exception -is [System.UnauthorizedAccessException]) -or ($msg -match 'denied|blocked|virus|defender|antivirus|security|Unauthorized')
    if ($isCorrupt) {
        Write-Host '[WARNING] Downloaded file on GitHub failed XML check; rejecting and falling back to local file.' -ForegroundColor Yellow
        Write-Host '[CANH BAO] File tai tu GitHub bi loi XML; tu choi ghi de va dung file hop le tren may.' -ForegroundColor Yellow
    } elseif ($isBlocked) {
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
        if (Test-Path -LiteralPath $tmpFile) { Remove-Item $tmpFile -Force -ErrorAction SilentlyContinue }
        exit 1
    }
    if (-not $localValid) {
        Write-Host '[ERROR] Local AcademicFormatter.dotm failed ZIP/XML integrity check. Please re-download from Releases.' -ForegroundColor Red
        Write-Host '[LOI] File AcademicFormatter.dotm tren may bi hong XML. Vui long tai lai tu Releases.' -ForegroundColor Red
        if (Test-Path -LiteralPath $tmpFile) { Remove-Item $tmpFile -Force -ErrorAction SilentlyContinue }
        exit 7
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
    if (Test-Path -LiteralPath $tmpFile) { Remove-Item $tmpFile -Force -ErrorAction SilentlyContinue }
    exit 2
}

if (-not $majorVer) {
    try { $majorVer = (Get-Item $wordExe).VersionInfo.FileMajorPart } catch { $majorVer = 16 }
}
Write-Host "Detected Microsoft Word v$majorVer ($wordExe)" -ForegroundColor Cyan

if ($majorVer -lt 15) {
    Write-Host '[ERROR] Word 2013 (v15) or newer is required. Your version is unsupported.' -ForegroundColor Red
    Write-Host '[LOI] Can Microsoft Word 2013 tro len. Phien ban hien tai khong ho tro.' -ForegroundColor Red
    if (Test-Path -LiteralPath $tmpFile) { Remove-Item $tmpFile -Force -ErrorAction SilentlyContinue }
    exit 3
}

$wordRunning = Get-Process -Name 'WINWORD' -ErrorAction SilentlyContinue
if ($wordRunning) {
    if (Test-Path -LiteralPath $tmpFile) { Remove-Item $tmpFile -Force -ErrorAction SilentlyContinue }
    Write-Host '[ERROR] Microsoft Word is currently running. Close ALL Word windows before installing/updating.' -ForegroundColor Red
    Write-Host '        (Prevents file lock and prevents Word from restoring Resiliency\DisabledItems on exit.)' -ForegroundColor Red
    Write-Host '[LOI] Microsoft Word dang mo. Hay dong TAT CA cua so Word truoc khi chay install-or-update.bat.' -ForegroundColor Red
    Write-Host '      (Tranh khoa file va tranh Word ghi de lai Resiliency\DisabledItems khi thoat.)' -ForegroundColor Red
    exit 4
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

$clearedPre = Clear-WordResiliencyBlock
try {
    Copy-Item -Path $installSrc -Destination $dest -Force
} catch {
    if ($tmpFile -and (Test-Path -LiteralPath $tmpFile)) { Remove-Item $tmpFile -Force -ErrorAction SilentlyContinue }
    Write-Host "[ERROR] Copy failed: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "[LOI] Khong the chep file: $($_.Exception.Message)" -ForegroundColor Red
    exit 5
}

try { Unblock-File -Path $dest -ErrorAction SilentlyContinue } catch {}
$clearedPost = Clear-WordResiliencyBlock
$totalCleared = $clearedPre + $clearedPost
if ($totalCleared -gt 0) {
    Write-Host "[OK] Cleared $totalCleared Word Resiliency\DisabledItems block(s) for AcademicFormatter.dotm." -ForegroundColor Cyan
    Write-Host "[OK] Da go bo $totalCleared muc chan trong Word Resiliency\DisabledItems." -ForegroundColor Cyan
}

$srcHash = (Get-FileHash $installSrc -Algorithm SHA256).Hash
$dstHash = (Get-FileHash $dest -Algorithm SHA256).Hash
if (Test-Path -LiteralPath $tmpFile) { Remove-Item $tmpFile -Force -ErrorAction SilentlyContinue }

if ($srcHash -ne $dstHash -or -not (Test-DotmIntegrity -FilePath $dest)) {
    Write-Host '[ERROR] File verification failed (SHA-256 or ZIP/XML integrity mismatch).' -ForegroundColor Red
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

Write-Host '[DONE] Open Word and select the Formatter Macro tab.' -ForegroundColor Green
Write-Host '[XONG] Mo Word va chon tab Formatter Macro de su dung.' -ForegroundColor Green
exit 0
# [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.