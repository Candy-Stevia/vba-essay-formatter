@echo off
REM [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
setlocal
title Academic Formatter Macro - Uninstaller (Go Cai Dat)
echo ============================================================================
echo  ACADEMIC FORMATTER MACRO - AUTO UNINSTALLER (GO CAI DAT TU DONG)
echo ============================================================================
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$c = (Get-Content -LiteralPath '%~f0' -Raw) -split ('#__' + 'PS_START__'), 2; Invoke-Expression $c[1]"
set "EC=%errorlevel%"
echo.
if not "%EC%"=="0" (
    echo Uninstallation stopped with code %EC%. ^(Go cai dat dung voi ma loi %EC%.^)
)
pause
exit /b %EC%

#__PS_START__
$ErrorActionPreference = 'Stop'

$dirs = New-Object System.Collections.Generic.List[string]
foreach ($v in @('16.0','15.0','14.0','12.0')) {
    $optKey = "HKCU:\Software\Microsoft\Office\$v\Word\Options"
    if (Test-Path $optKey) {
        $custom = (Get-ItemProperty $optKey -Name 'STARTUP-PATH' -ErrorAction SilentlyContinue).'STARTUP-PATH'
        if ($custom -and $custom.Trim().Length -gt 0) {
            $expanded = [Environment]::ExpandEnvironmentVariables($custom.Trim())
            if (-not $dirs.Contains($expanded)) { $dirs.Add($expanded) }
        }
    }
}

$defaultStartup = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP'
if (-not $dirs.Contains($defaultStartup)) { $dirs.Add($defaultStartup) }

$targets = New-Object System.Collections.Generic.List[string]
foreach ($d in $dirs) {
    foreach ($name in @('AcademicFormatter.dotm', 'AcademicFormatter.dotm.bak')) {
        $p = Join-Path $d $name
        if (Test-Path -LiteralPath $p) {
            $targets.Add($p)
        }
    }
}

if ($targets.Count -eq 0) {
    Write-Host '[INFO] AcademicFormatter.dotm is not installed (already removed).' -ForegroundColor Cyan
    Write-Host '[THONG BAO] Khong tim thay AcademicFormatter.dotm (da duoc go cai dat).' -ForegroundColor Cyan
    exit 0
}

$wordRunning = Get-Process -Name 'WINWORD' -ErrorAction SilentlyContinue
$removedCount = 0

foreach ($t in $targets) {
    try {
        Remove-Item -LiteralPath $t -Force -ErrorAction Stop
        if (Test-Path -LiteralPath $t) {
            throw "File still exists after deletion."
        }
        Write-Host "[OK] Removed: $t" -ForegroundColor Green
        Write-Host "[OK] Da xoa:  $t" -ForegroundColor Green
        $removedCount++
    } catch {
        if ($wordRunning) {
            Write-Host '[ERROR] Word is open and locking AcademicFormatter.dotm. Close Word and run again.' -ForegroundColor Red
            Write-Host '[LOI] Word dang mo va khoa file. Hay dong Word roi chay lai Uninstall.bat.' -ForegroundColor Red
            exit 4
        }
        Write-Host "[ERROR] Failed to remove $t : $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "[LOI] Khong the xoa file $t : $($_.Exception.Message)" -ForegroundColor Red
        exit 5
    }
}

if ($wordRunning) {
    Write-Host '[NOTE] Word is currently open. Restart Word to unload the Formatter Macro tab.' -ForegroundColor Yellow
    Write-Host '[LUU Y] Word dang mo. Hay dong va mo lai Word de bo tab Formatter Macro.' -ForegroundColor Yellow
} else {
    foreach ($ov in @('16.0', '15.0', '14.0', '12.0')) {
        foreach ($sub in @('DisabledItems', 'StartupItems')) {
            $rk = "HKCU:\Software\Microsoft\Office\$ov\Word\Resiliency\$sub"
            if (Test-Path $rk) {
                $props = Get-ItemProperty $rk -ErrorAction SilentlyContinue
                foreach ($pr in $props.PSObject.Properties) {
                    if ($pr.Name -notlike 'PS*') {
                        try {
                            $txt = [System.Text.Encoding]::Unicode.GetString([byte[]]$pr.Value) -replace "`0", ''
                            if ($txt -match 'AcademicFormatter\.dotm') { Remove-ItemProperty -Path $rk -Name $pr.Name -ErrorAction SilentlyContinue }
                        } catch {}
                    }
                }
            }
        }
    }
    Write-Host '[DONE] Academic Formatter Macro has been completely uninstalled.' -ForegroundColor Green
    Write-Host '[XONG] Da go cai dat hoan toan Academic Formatter Macro.' -ForegroundColor Green
}
exit 0
# [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
