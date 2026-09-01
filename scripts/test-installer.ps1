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

if (Get-Process ExcelDiffTracker -ErrorAction SilentlyContinue) {
    throw 'Excel Diff Tracker is already running. Use a clean test VM or Windows account.'
}
if (Test-Path $installDirectory) {
    throw 'An existing Excel Diff Tracker installation was found. This test installs and uninstalls the app, so use a clean test VM or Windows account.'
}

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

$testDataDirectory = Join-Path $env:TEMP "ExcelDiffTracker-InstallerSmoke-$([Guid]::NewGuid().ToString('N'))"
$previousTestDataDirectory = $env:EXCEL_DIFF_TRACKER_TEST_DATA_DIRECTORY
$applicationProcess = $null
try {
    $env:EXCEL_DIFF_TRACKER_TEST_DATA_DIRECTORY = $testDataDirectory
    $applicationProcess = Start-Process -FilePath $application -PassThru
    $deadline = [DateTime]::UtcNow.AddSeconds(15)
    do {
        Start-Sleep -Milliseconds 250
        $applicationProcess.Refresh()
        if ($applicationProcess.HasExited) {
            throw "Installed application exited with code $($applicationProcess.ExitCode)."
        }
        $startupTitle = $applicationProcess.MainWindowTitle
    } while ([string]::IsNullOrWhiteSpace($startupTitle) -and [DateTime]::UtcNow -lt $deadline)

    if ($startupTitle -ne 'Welcome to Excel Diff Tracker') {
        throw "First-run onboarding did not open successfully. Visible window: '$startupTitle'."
    }
}
finally {
    $env:EXCEL_DIFF_TRACKER_TEST_DATA_DIRECTORY = $previousTestDataDirectory
    if ($applicationProcess -and -not $applicationProcess.HasExited) {
        $applicationProcess.Kill()
        $applicationProcess.WaitForExit()
    }
    if (Test-Path $testDataDirectory) {
        Remove-Item $testDataDirectory -Recurse -Force
    }
}

$uninstaller = Join-Path $installDirectory 'unins000.exe'
if (-not (Test-Path $uninstaller)) { throw 'Uninstaller was not found.' }
$uninstallProcess = Start-Process -FilePath $uninstaller -ArgumentList '/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART' -Wait -PassThru
if ($uninstallProcess.ExitCode -ne 0) { throw "Uninstaller exited with code $($uninstallProcess.ExitCode)." }
if (Test-Path $installDirectory) { throw 'The application directory remained after uninstall.' }
if (Test-Path $startMenuShortcut) { throw 'The Start-menu shortcut remained after uninstall.' }

$remainingRunValue = (Get-ItemProperty -Path $runKey -Name ExcelDiffTracker -ErrorAction SilentlyContinue).ExcelDiffTracker
if ($remainingRunValue) { throw 'The startup registration remained after uninstall.' }

Write-Host 'INSTALLER_SMOKE_PASS'
