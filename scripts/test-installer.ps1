[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$InstallerPath,
    [switch]$RequireNoDotnet
)

$ErrorActionPreference = 'Stop'
$installer = (Resolve-Path $InstallerPath).Path
$installDirectory = Join-Path $env:LOCALAPPDATA 'Programs\Excel Diff Tracker'
$application = Join-Path $installDirectory 'ExcelDiffTracker.exe'
$runKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'
$startMenuShortcut = Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs\Excel Diff Tracker\Excel Diff Tracker.lnk'
$windowsSdkLicense = Join-Path $installDirectory 'licenses\dotnet\Microsoft-Windows-SDK-10.0.26100.57-License.rtf'
$expectedWindowsSdkLicenseHash = 'DD07EB178E00C6BBA4148457FC00FF77CD4887EB521D504186FE59C9EC8BBE62'

if ($RequireNoDotnet -and (Get-Command dotnet -ErrorAction SilentlyContinue)) {
    throw 'dotnet is available on PATH; this run cannot prove a no-runtime installation.'
}

Get-Process ExcelDiffTracker -ErrorAction SilentlyContinue | Stop-Process -Force
$installProcess = Start-Process -FilePath $installer -ArgumentList '/VERYSILENT', '/CURRENTUSER', '/SUPPRESSMSGBOXES', '/NORESTART', '/TASKS="startup"' -Wait -PassThru
if ($installProcess.ExitCode -ne 0) { throw "Installer exited with code $($installProcess.ExitCode)." }
if (-not (Test-Path $application)) { throw 'Installed application was not found.' }
if (-not (Test-Path $startMenuShortcut)) { throw 'Start-menu shortcut was not found.' }
if (-not (Test-Path $windowsSdkLicense)) { throw 'The installed Windows SDK license was not found.' }
if ((Get-FileHash $windowsSdkLicense -Algorithm SHA256).Hash -ne $expectedWindowsSdkLicenseHash) {
    throw 'The installed Windows SDK license did not match the approved source.'
}

$runValue = (Get-ItemProperty -Path $runKey -Name ExcelDiffTracker -ErrorAction Stop).ExcelDiffTracker
if ($runValue -notmatch 'ExcelDiffTracker\.exe') { throw 'Startup registration is missing or invalid.' }

$applicationProcess = Start-Process -FilePath $application -ArgumentList '--background' -PassThru
Start-Sleep -Seconds 5
if ($applicationProcess.HasExited) { throw "Installed application exited with code $($applicationProcess.ExitCode)." }
$applicationProcess.Kill()
$applicationProcess.WaitForExit()

$uninstaller = Join-Path $installDirectory 'unins000.exe'
if (-not (Test-Path $uninstaller)) { throw 'Uninstaller was not found.' }
$uninstallProcess = Start-Process -FilePath $uninstaller -ArgumentList '/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART' -Wait -PassThru
if ($uninstallProcess.ExitCode -ne 0) { throw "Uninstaller exited with code $($uninstallProcess.ExitCode)." }
if (Test-Path $installDirectory) { throw 'The application directory remained after uninstall.' }
if (Test-Path $startMenuShortcut) { throw 'The Start-menu shortcut remained after uninstall.' }

$remainingRunValue = (Get-ItemProperty -Path $runKey -Name ExcelDiffTracker -ErrorAction SilentlyContinue).ExcelDiffTracker
if ($remainingRunValue) { throw 'The startup registration remained after uninstall.' }

Write-Host 'INSTALLER_SMOKE_PASS'
