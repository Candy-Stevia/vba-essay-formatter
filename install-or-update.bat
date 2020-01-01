@echo off
REM [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
setlocal
title Academic Formatter Macro v1.0.8 (Admonition Update) - Install or Update
echo ============================================================================
echo  ACADEMIC FORMATTER MACRO v1.0.8 (ADMONITION UPDATE) - INSTALL OR UPDATE
echo ============================================================================
echo.

set "SRC_DOTM=%~dp0AcademicFormatter.dotm"
set "SELF_BAT=%~f0"
set "INSTALL_ARGS=%*"

powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$c = (Get-Content -LiteralPath $env:SELF_BAT -Raw) -split ('#__' + 'PS_START__'), 2; Invoke-Expression $c[1]"
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
    if (-not $fi -or $fi.Length -lt 50000 -or $fi.Length -gt 10485760) { return $false }
    $hdr = New-Object byte[] 2
    $fs = $null
    try {
        $fs = [System.IO.File]::Open($FilePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        if ($fs.Read($hdr, 0, 2) -lt 2) { return $false }
    } catch {
        return $false
    } finally {
        if ($fs) { $fs.Dispose() }
    }
    if ($hdr[0] -ne 0x50 -or $hdr[1] -ne 0x4B) { return $false }
    $zip = $null
    try {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($FilePath)
        $coreEntry = $zip.GetEntry('docProps/core.xml')
        $uiEntry = $zip.GetEntry('customUI/customUI14.xml')
        $vbaEntry = $zip.GetEntry('word/vbaProject.bin')
        if (-not $coreEntry -or -not $uiEntry -or -not $vbaEntry) { return $false }
        foreach ($entry in @($coreEntry, $uiEntry)) {
            $sr = New-Object System.IO.StreamReader($entry.Open())
            try { $xmlText = $sr.ReadToEnd() } finally { $sr.Dispose() }
            $docXml = New-Object System.Xml.XmlDocument
            $docXml.XmlResolver = $null
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
            if (Test-Path -LiteralPath $regKey) {
                $props = Get-ItemProperty -LiteralPath $regKey -ErrorAction SilentlyContinue
                foreach ($pr in $props.PSObject.Properties) {
                    if ($pr.Name -notlike 'PS*') {
                        try {
                            $txt = [System.Text.Encoding]::Unicode.GetString([byte[]]$pr.Value) -replace "`0", ''
                            if ($txt -match 'AcademicFormatter\.dotm') {
                                Remove-ItemProperty -LiteralPath $regKey -Name $pr.Name -ErrorAction SilentlyContinue
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

function Test-InstructionContentSignature {
    param([string]$HeadText)
    if (-not $HeadText) { return $false }
    if ($HeadText -match '<!--\s*AF-INSTRUCTION-SIGNATURE:\s*AcademicFormatter\s*-->') { return $true }
    if (($HeadText -match 'Academic Document Format Engine') -and ($HeadText -match 'AcademicFormatter\.dotm')) { return $true }
    return $false
}

function Test-InstructionIntegrity {
    param([string]$FilePath)
    if (-not (Test-Path -LiteralPath $FilePath)) { return $false }
    $fi = Get-Item -LiteralPath $FilePath -ErrorAction SilentlyContinue
    if (-not $fi -or $fi.Length -lt 3072 -or $fi.Length -gt 512000) { return $false }
    try {
        $raw = [System.IO.File]::ReadAllText($FilePath, [System.Text.Encoding]::UTF8)
        if (-not $raw) { return $false }
        $trimHead = $raw.TrimStart([char]0xFEFF, ' ', "`t", "`r", "`n")
        if ($trimHead -imatch '^(<!DOCTYPE|<html)') { return $false }
        $head = if ($raw.Length -gt 4096) { $raw.Substring(0, 4096) } else { $raw }
        return (Test-InstructionContentSignature -HeadText $head)
    } catch {
        return $false
    }
}

function Resolve-InstructionTarget {
    param([string]$ScriptDir)
    $defaultPath = Join-Path $ScriptDir 'instruction-108.md'
    $altPath = Join-Path $ScriptDir 'AcademicFormatter_instruction-108.md'
    if (-not (Test-Path -LiteralPath $defaultPath)) {
        return [pscustomobject]@{ TargetPath = $defaultPath; IsForeignConflict = $false }
    }
    try {
        $sr = New-Object System.IO.StreamReader($defaultPath, [System.Text.Encoding]::UTF8, $true)
        try {
            $buf = New-Object char[] 4096
            $readLen = $sr.ReadBlock($buf, 0, 4096)
            $head = New-Object string($buf, 0, $readLen)
        } finally {
            $sr.Dispose()
        }
        if (Test-InstructionContentSignature -HeadText $head) {
            return [pscustomobject]@{ TargetPath = $defaultPath; IsForeignConflict = $false }
        }
    } catch {}
    return [pscustomobject]@{ TargetPath = $altPath; IsForeignConflict = $true }
}

function Sync-InstructionManual {
    param(
        [string]$ScriptDir,
        [string[]]$Urls = @(
            'https://github.com/Candy-Stevia/vba-essay-formatter/releases/latest/download/instruction-108.md',
            'https://raw.githubusercontent.com/Candy-Stevia/vba-essay-formatter/main/instruction-108.md'
        )
    )
    $res = Resolve-InstructionTarget -ScriptDir $ScriptDir
    $targetPath = $res.TargetPath
    if ($res.IsForeignConflict) {
        Write-Host '[WARNING] Existing instruction-108.md does not belong to AcademicFormatter; keeping it untouched.' -ForegroundColor Yellow
        Write-Host '          Saving AcademicFormatter guide to: AcademicFormatter_instruction-108.md' -ForegroundColor Yellow
        Write-Host '[CANH BAO] instruction-108.md hien co thuoc du an khac; giu nguyen va luu vao AcademicFormatter_instruction-108.md.' -ForegroundColor Yellow
    }

    $tmpInst = Join-Path ([System.IO.Path]::GetTempPath()) ('AF_inst_' + [guid]::NewGuid().ToString('N') + '.tmp')
    $hasTarget = Test-Path -LiteralPath $targetPath
    $status = 'kept_local'
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $dlOk = $false
        $lastEx = $null
        foreach ($u in $Urls) {
            try {
                Invoke-WebRequest -Uri $u -OutFile $tmpInst -TimeoutSec 15 -UseBasicParsing
                $dlOk = $true
                break
            } catch {
                $lastEx = $_
                continue
            }
        }
        if (-not $dlOk -and $lastEx) { throw $lastEx }
        if (-not (Test-InstructionIntegrity -FilePath $tmpInst)) {
            throw [System.InvalidOperationException]::new('Downloaded instruction-108.md failed integrity or signature check.')
        }
        $remoteHash = (Get-FileHash -LiteralPath $tmpInst -Algorithm SHA256).Hash
        $localHash = $null
        if ($hasTarget) { $localHash = (Get-FileHash -LiteralPath $targetPath -Algorithm SHA256).Hash }
        if ($localHash -and $remoteHash -eq $localHash) {
            Remove-Item -LiteralPath $tmpInst -Force -ErrorAction SilentlyContinue
            $status = 'up_to_date'
            Write-Host '[UP TO DATE] instruction-108.md is already up to date.' -ForegroundColor Green
            Write-Host '[MOI NHAT] Huong dan instruction-108.md da la ban moi nhat.' -ForegroundColor Green
        } else {
            Move-Item -LiteralPath $tmpInst -Destination $targetPath -Force
            try { Unblock-File -LiteralPath $targetPath -ErrorAction SilentlyContinue } catch {}
            $status = 'updated'
            Write-Host "[OK] Synchronized latest guide to: $targetPath" -ForegroundColor Green
            Write-Host "[OK] Da cap nhat huong dan moi nhat tai: $targetPath" -ForegroundColor Green
        }
    } catch {
        if (Test-Path -LiteralPath $tmpInst) { Remove-Item -LiteralPath $tmpInst -Force -ErrorAction SilentlyContinue }
        $status = if ($hasTarget) { 'offline_kept' } else { 'offline_missing' }
        Write-Host "[WARNING] Could not auto-update guide online ($($_.Exception.Message)); continuing safely." -ForegroundColor Yellow
        Write-Host '[CANH BAO] Khong cap nhat duoc huong dan online (offline hoac file bi khoa); giu nguyen ban hien co.' -ForegroundColor Yellow
    }

    $fullPath = if (Test-Path -LiteralPath $targetPath) { [System.IO.Path]::GetFullPath($targetPath) } else { '' }
    return [pscustomobject]@{
        Status = $status
        TargetPath = $targetPath
        FullPath = $fullPath
        IsForeignConflict = $res.IsForeignConflict
    }
}

#__PS_MAIN__

$wordExe = $null
$majorVer = $null

foreach ($r in @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe')) {
    if (Test-Path -LiteralPath $r) {
        $p = (Get-ItemProperty -LiteralPath $r -ErrorAction SilentlyContinue).'(default)'
        if ($p -and (Test-Path -LiteralPath $p)) { $wordExe = $p; break }
    }
}

if (-not $wordExe) {
    foreach ($v in @('16.0','15.0','14.0','12.0')) {
        foreach ($base in @('HKLM:\SOFTWARE\Microsoft\Office', 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office')) {
            $k = "$base\$v\Word\InstallRoot"
            if (Test-Path -LiteralPath $k) {
                $dir = (Get-ItemProperty -LiteralPath $k -ErrorAction SilentlyContinue).Path
                if ($dir -and (Test-Path -LiteralPath (Join-Path $dir 'WINWORD.EXE'))) {
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
    try { $majorVer = (Get-Item -LiteralPath $wordExe).VersionInfo.FileMajorPart } catch { $majorVer = 16 }
}
Write-Host "Detected Microsoft Word v$majorVer ($wordExe)" -ForegroundColor Cyan

if ($majorVer -lt 15) {
    Write-Host '[ERROR] Word 2013 (v15) or newer is required. Your version is unsupported.' -ForegroundColor Red
    Write-Host '[LOI] Can Microsoft Word 2013 tro len. Phien ban hien tai khong ho tro.' -ForegroundColor Red
    exit 3
}

$wordRunning = Get-Process -Name 'WINWORD' -ErrorAction SilentlyContinue
if ($wordRunning) {
    Write-Host '[ERROR] Microsoft Word is currently running. Close ALL Word windows before installing/updating.' -ForegroundColor Red
    Write-Host '        (Prevents file lock and prevents Word from restoring Resiliency\DisabledItems on exit.)' -ForegroundColor Red
    Write-Host '[LOI] Microsoft Word dang mo. Hay dong TAT CA cua so Word truoc khi chay install-or-update.bat.' -ForegroundColor Red
    Write-Host '      (Tranh khoa file va tranh Word ghi de lai Resiliency\DisabledItems khi thoat.)' -ForegroundColor Red
    exit 4
}

$localSrc = $env:SRC_DOTM
$hasLocal = Test-Path -LiteralPath $localSrc
$localValid = $hasLocal -and (Test-DotmIntegrity -FilePath $localSrc)
$installSrc = $localSrc
$isOffline = ($env:INSTALL_ARGS -match '(?i)(^|\s)(-Offline|--offline)(\s|$)')

$releaseUrl = 'https://github.com/Candy-Stevia/vba-essay-formatter/releases/latest/download/AcademicFormatter.dotm'
$rawMainUrl = 'https://raw.githubusercontent.com/Candy-Stevia/vba-essay-formatter/main/AcademicFormatter.dotm'
$manualUrl = 'https://github.com/Candy-Stevia/vba-essay-formatter/releases/latest'

$tmpFile = Join-Path $env:TEMP ('AF_update_' + [guid]::NewGuid().ToString('N') + '.dotm')
$downloadedOk = $false

if ($isOffline) {
    Write-Host '[OFFLINE] Offline mode requested. Using local AcademicFormatter.dotm on this PC.' -ForegroundColor Cyan
    Write-Host '[NGOAI TUYEN] Che do offline. Dang dung file AcademicFormatter.dotm co san tren may.' -ForegroundColor Cyan
    if (-not $hasLocal) {
        Write-Host '[ERROR] No local AcademicFormatter.dotm found next to install-or-update.bat.' -ForegroundColor Red
        Write-Host '[LOI] Khong tim thay file AcademicFormatter.dotm co san cung thu muc.' -ForegroundColor Red
        exit 1
    }
    if (-not $localValid) {
        Write-Host '[ERROR] Local AcademicFormatter.dotm failed ZIP/XML integrity check. Please re-download from Releases.' -ForegroundColor Red
        Write-Host '[LOI] File AcademicFormatter.dotm tren may bi hong XML. Vui long tai lai tu Releases.' -ForegroundColor Red
        exit 7
    }
    $installSrc = $localSrc
} else {
    Write-Host '[CHECK] Checking GitHub for latest version... (Dang kiem tra ban moi tren GitHub...)' -ForegroundColor Cyan

    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $urls = @($releaseUrl, $rawMainUrl)
        $lastEx = $null
        $dlResp = $null
        foreach ($u in $urls) {
            try {
                $dlResp = Invoke-WebRequest -Uri $u -OutFile $tmpFile -PassThru -TimeoutSec 15 -UseBasicParsing
                $downloadedOk = $true
                break
            } catch {
                $lastEx = $_
                continue
            }
        }
        if (-not $downloadedOk -and $lastEx) { throw $lastEx }

        try { Unblock-File -LiteralPath $tmpFile -ErrorAction SilentlyContinue } catch {}
        if (-not (Test-Path -LiteralPath $tmpFile)) {
            throw [System.UnauthorizedAccessException]::new('Downloaded file was removed or blocked by security software.')
        }
        $tmpInfo = Get-Item -LiteralPath $tmpFile -ErrorAction SilentlyContinue
        if (-not $tmpInfo -or $tmpInfo.Length -gt 10485760) {
            throw [System.InvalidOperationException]::new('CORRUPT_XML: Downloaded .dotm exceeds 10 MB safety limit.')
        }
        if (-not (Test-DotmIntegrity -FilePath $tmpFile)) {
            throw [System.InvalidOperationException]::new('CORRUPT_XML: Downloaded .dotm failed ZIP/XML integrity verification.')
        }

        $remoteLastMod = $null
        if ($dlResp -and $dlResp.Headers -and $dlResp.Headers['Last-Modified']) {
            $parsedDt = [datetime]::MinValue
            if ([datetime]::TryParse($dlResp.Headers['Last-Modified'], [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AdjustToUniversal, [ref]$parsedDt)) {
                $remoteLastMod = $parsedDt
            }
        }

        $remoteHash = (Get-FileHash -LiteralPath $tmpFile -Algorithm SHA256).Hash
        $localHash = $null
        $localLastMod = $null
        if ($localValid) {
            $localHash = (Get-FileHash -LiteralPath $localSrc -Algorithm SHA256).Hash
            $localLastMod = (Get-Item -LiteralPath $localSrc -ErrorAction SilentlyContinue).LastWriteTimeUtc
        }

        if ($localHash -and $remoteHash -eq $localHash) {
            Write-Host '[UP TO DATE] You already have the latest version on this PC.' -ForegroundColor Green
            Write-Host '[MOI NHAT] Ban dang su dung phien ban moi nhat co san tren may.' -ForegroundColor Green
            $installSrc = $localSrc
        } elseif ($localValid -and $localLastMod -and $remoteLastMod -and ($localLastMod -ge $remoteLastMod)) {
            Write-Host '[LOCAL NEWER] Local AcademicFormatter.dotm is newer than GitHub; keeping local version.' -ForegroundColor Cyan
            Write-Host '[BAN LOCAL MOI HON] File AcademicFormatter.dotm tren may moi hon ban GitHub; giu nguyen ban tren may.' -ForegroundColor Cyan
            $installSrc = $localSrc
        } else {
            Write-Host '[UPDATE] Newer valid version found on GitHub! Using downloaded AcademicFormatter.dotm.' -ForegroundColor Green
            Write-Host '[CAP NHAT] Da co phien ban moi hop le tren GitHub! Dang nap ban moi.' -ForegroundColor Green
            $installSrc = $tmpFile
            if ((-not $hasLocal) -or (-not $localValid) -or ($remoteLastMod -and $localLastMod -and ($remoteLastMod -gt $localLastMod))) {
                try { Copy-Item -LiteralPath $tmpFile -Destination $localSrc -Force -ErrorAction SilentlyContinue } catch {}
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
            if (Test-Path -LiteralPath $tmpFile) { Remove-Item -LiteralPath $tmpFile -Force -ErrorAction SilentlyContinue }
            exit 1
        }
        if (-not $localValid) {
            Write-Host '[ERROR] Local AcademicFormatter.dotm failed ZIP/XML integrity check. Please re-download from Releases.' -ForegroundColor Red
            Write-Host '[LOI] File AcademicFormatter.dotm tren may bi hong XML. Vui long tai lai tu Releases.' -ForegroundColor Red
            if (Test-Path -LiteralPath $tmpFile) { Remove-Item -LiteralPath $tmpFile -Force -ErrorAction SilentlyContinue }
            exit 7
        }
        $installSrc = $localSrc
    }
}

$verKey = "$majorVer.0"
if ($majorVer -ge 16) { $verKey = '16.0' }

$startupDir = $null
$optKey = "HKCU:\Software\Microsoft\Office\$verKey\Word\Options"
if (Test-Path -LiteralPath $optKey) {
    $custom = (Get-ItemProperty -LiteralPath $optKey -Name 'STARTUP-PATH' -ErrorAction SilentlyContinue).'STARTUP-PATH'
    if ($custom -and $custom.Trim().Length -gt 0) {
        $startupDir = [Environment]::ExpandEnvironmentVariables($custom.Trim())
    }
}
if (-not $startupDir) {
    $startupDir = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP'
}
if (-not (Test-Path -LiteralPath $startupDir)) {
    New-Item -ItemType Directory -Path $startupDir -Force | Out-Null
}

$dest = Join-Path $startupDir 'AcademicFormatter.dotm'

$clearedPre = Clear-WordResiliencyBlock
try {
    Copy-Item -LiteralPath $installSrc -Destination $dest -Force
} catch {
    if ($tmpFile -and (Test-Path -LiteralPath $tmpFile)) { Remove-Item -LiteralPath $tmpFile -Force -ErrorAction SilentlyContinue }
    Write-Host "[ERROR] Copy failed: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "[LOI] Khong the chep file: $($_.Exception.Message)" -ForegroundColor Red
    exit 5
}

try { Unblock-File -LiteralPath $dest -ErrorAction SilentlyContinue } catch {}
$clearedPost = Clear-WordResiliencyBlock
$totalCleared = $clearedPre + $clearedPost
if ($totalCleared -gt 0) {
    Write-Host "[OK] Cleared $totalCleared Word Resiliency\DisabledItems block(s) for AcademicFormatter.dotm." -ForegroundColor Cyan
    Write-Host "[OK] Da go bo $totalCleared muc chan trong Word Resiliency\DisabledItems." -ForegroundColor Cyan
}

$srcHash = $null
$dstHash = $null
try {
    if (-not (Test-Path -LiteralPath $dest)) {
        throw [System.IO.FileNotFoundException]::new('Destination file missing after copy (possibly quarantined by Windows Defender).', $dest)
    }
    $srcHash = (Get-FileHash -LiteralPath $installSrc -Algorithm SHA256).Hash
    $dstHash = (Get-FileHash -LiteralPath $dest -Algorithm SHA256).Hash
} catch {
    if ($tmpFile -and (Test-Path -LiteralPath $tmpFile)) { Remove-Item -LiteralPath $tmpFile -Force -ErrorAction SilentlyContinue }
    Write-Host "[ERROR] File verification failed (possibly quarantined by Windows Defender): $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "[LOI] Kiem tra file that bai (co the bi Windows Defender cach ly): $($_.Exception.Message)" -ForegroundColor Red
    exit 6
}
if ($tmpFile -and (Test-Path -LiteralPath $tmpFile)) { Remove-Item -LiteralPath $tmpFile -Force -ErrorAction SilentlyContinue }

if ($srcHash -ne $dstHash -or -not (Test-DotmIntegrity -FilePath $dest)) {
    Write-Host '[ERROR] File verification failed (SHA-256 or ZIP/XML integrity mismatch).' -ForegroundColor Red
    Write-Host '[LOI] Kiem tra toan ven file that bai.' -ForegroundColor Red
    exit 6
}

Write-Host "[OK] Installed to: $dest" -ForegroundColor Green
Write-Host "[OK] Da cai dat vao: $dest" -ForegroundColor Green

$scriptDir = Split-Path -Parent $env:SELF_BAT
if (-not $scriptDir) { $scriptDir = (Get-Location).Path }
$instSync = Sync-InstructionManual -ScriptDir $scriptDir

$vbaWarn = $null
foreach ($sk in @("HKLM:\Software\Policies\Microsoft\Office\$verKey\Word\Security", "HKCU:\Software\Policies\Microsoft\Office\$verKey\Word\Security", "HKCU:\Software\Microsoft\Office\$verKey\Word\Security")) {
    if (Test-Path -LiteralPath $sk) {
        $w = (Get-ItemProperty -LiteralPath $sk -Name 'VBAWarnings' -ErrorAction SilentlyContinue).VBAWarnings
        if ($null -ne $w) { $vbaWarn = [int]$w; break }
    }
}
if ($vbaWarn -eq 4) {
    Write-Host '[WARNING] Word Macro Settings are set to Disable all without notification.' -ForegroundColor Yellow
    Write-Host '          Open Word > File > Options > Trust Center > Macro Settings to enable.' -ForegroundColor Yellow
    Write-Host '[CANH BAO] Word dang tat tat ca Macro khong thong bao.' -ForegroundColor Yellow
    Write-Host '           Vao File > Options > Trust Center > Macro Settings de bat.' -ForegroundColor Yellow
}

if ($instSync -and $instSync.FullPath) {
    Write-Host ''
    Write-Host '============================================================================' -ForegroundColor Cyan
    Write-Host "[AI GUIDE] Latest manual ready at: $($instSync.FullPath)" -ForegroundColor Cyan
    Write-Host '           Always send THIS exact newest file to AI (ChatGPT/Claude/Gemini).' -ForegroundColor White
    Write-Host "[HUONG DAN AI] Ban moi nhat dang o: $($instSync.FullPath)" -ForegroundColor Cyan
    Write-Host '               Hay gui DUNG file tai dia chi nay cho AI de tranh sai phien ban!' -ForegroundColor White
    Write-Host '============================================================================' -ForegroundColor Cyan
}

Write-Host '[DONE] Open Word and select the Formatter Macro tab.' -ForegroundColor Green
Write-Host '[XONG] Mo Word va chon tab Formatter Macro de su dung.' -ForegroundColor Green
exit 0
# [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.