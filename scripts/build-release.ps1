[CmdletBinding()]
param(
    [ValidatePattern('^\d+\.\d+\.\d+([-.][0-9A-Za-z.-]+)?$')]
    [string]$Version = '0.1.2',
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Release',
    [switch]$SkipInstaller
)

$ErrorActionPreference = 'Stop'
$repositoryRoot = Split-Path -Parent $PSScriptRoot
$solution = Join-Path $repositoryRoot 'ExcelDiffTracker.slnx'
$artifacts = Join-Path $repositoryRoot 'artifacts'
$publish = Join-Path $artifacts 'publish\win-arm64'
$acceptanceTools = Join-Path $artifacts 'acceptance-tools\win-arm64'
$smokeTool = Join-Path $artifacts 'smoke-tool'
$branding = Join-Path $artifacts 'branding'
$release = Join-Path $artifacts 'release'
$icon = Join-Path $branding 'app-icon.ico'
$wingetManifestRelative = "packaging/winget/yogoldy.ExcelDiffTracker/$Version/yogoldy.ExcelDiffTracker.installer.yaml"
$wingetManifest = Join-Path $repositoryRoot $wingetManifestRelative.Replace('/', [System.IO.Path]::DirectorySeparatorChar)

$sourceCommit = (& git -C $repositoryRoot rev-parse HEAD).Trim().ToLowerInvariant()
if ($LASTEXITCODE -ne 0 -or $sourceCommit -notmatch '^[a-f0-9]{40}$') {
    throw 'A frozen Git source commit is required before building a release.'
}
$dirtySource = @(& git -C $repositoryRoot status --porcelain --untracked-files=all)
if ($LASTEXITCODE -ne 0 -or $dirtySource.Count -ne 0) {
    throw 'The source tree must be clean before building a release candidate.'
}
if (-not (Test-Path $wingetManifest -PathType Leaf)) {
    throw "The candidate WinGet manifest is missing: $wingetManifest"
}

function Get-SourceManifestContent {
    $trackedFiles = @(& git -C $repositoryRoot ls-files)
    if ($LASTEXITCODE -ne 0 -or $trackedFiles.Count -eq 0) { throw 'Could not enumerate tracked source files.' }
    $includedFiles = @($trackedFiles | Where-Object { $_ -ne $wingetManifestRelative } | Sort-Object)
    if ($includedFiles.Count -ne ($trackedFiles.Count - 1)) {
        throw 'The generated WinGet hash manifest must be the one and only source-manifest exclusion.'
    }
    $lines = foreach ($relative in $includedFiles) {
        $fullPath = Join-Path $repositoryRoot $relative.Replace('/', [System.IO.Path]::DirectorySeparatorChar)
        if (-not (Test-Path $fullPath -PathType Leaf)) { throw "Tracked source file is missing: $relative" }
        $hash = (Get-FileHash -LiteralPath $fullPath -Algorithm SHA256).Hash.ToLowerInvariant()
        "$hash  $relative"
    }
    [pscustomobject]@{ Text = (($lines -join "`n") + "`n"); FileCount = $includedFiles.Count }
}

$sourceManifest = Get-SourceManifestContent

if (Test-Path $artifacts) {
    Remove-Item $artifacts -Recurse -Force
}
New-Item -ItemType Directory -Path $publish, $acceptanceTools, $smokeTool, $branding, $release -Force | Out-Null

& (Join-Path $PSScriptRoot 'generate-icon.ps1') -OutputPath $icon

dotnet restore $solution
if ($LASTEXITCODE -ne 0) { throw 'Restore failed.' }
dotnet build $solution -c $Configuration --no-restore
if ($LASTEXITCODE -ne 0) { throw 'Build failed.' }
dotnet test $solution -c $Configuration --no-build
if ($LASTEXITCODE -ne 0) { throw 'Tests failed.' }

$appProject = Join-Path $repositoryRoot 'src\ExcelDiffTracker.App\ExcelDiffTracker.App.csproj'
dotnet publish $appProject -c $Configuration -r win-arm64 --self-contained true --no-restore -o $publish "-p:Version=$Version" "-p:ApplicationIcon=$icon"
if ($LASTEXITCODE -ne 0) { throw 'ARM64 publish failed.' }

$sourceManifestPath = Join-Path $publish 'BUILD-SOURCE-SHA256SUMS.txt'
[System.IO.File]::WriteAllText($sourceManifestPath, $sourceManifest.Text, [System.Text.UTF8Encoding]::new($false))
$sourceManifestHash = (Get-FileHash -LiteralPath $sourceManifestPath -Algorithm SHA256).Hash.ToUpperInvariant()
$buildIdentity = [ordered]@{
    schemaVersion = 1
    product = 'Excel Diff Tracker'
    version = $Version
    sourceCommit = $sourceCommit
    sourceManifest = 'BUILD-SOURCE-SHA256SUMS.txt'
    sourceManifestSha256 = $sourceManifestHash
    sourceFileCount = $sourceManifest.FileCount
    excludedGeneratedPath = $wingetManifestRelative
    createdUtc = [DateTime]::UtcNow.ToString('O')
}
[System.IO.File]::WriteAllText(
    (Join-Path $publish 'BUILD-IDENTITY.json'),
    ($buildIdentity | ConvertTo-Json -Depth 5),
    [System.Text.UTF8Encoding]::new($false))

$acceptanceProbeProject = Join-Path $repositoryRoot 'tools\ExcelDiffTracker.AcceptanceProbe\ExcelDiffTracker.AcceptanceProbe.csproj'
dotnet publish $acceptanceProbeProject -c $Configuration -r win-arm64 --self-contained true --no-restore -o $acceptanceTools
if ($LASTEXITCODE -ne 0) { throw 'ARM64 acceptance-probe publish failed.' }

# The frozen EXE hashes must cover all executable product/probe code, including
# managed assemblies and native runtime libraries. Symbols and notices may remain external.
foreach ($payload in @(
    @{ Directory = $publish; Executable = 'ExcelDiffTracker.exe' },
    @{ Directory = $acceptanceTools; Executable = 'ExcelDiffTracker.AcceptanceProbe.exe' }
)) {
    if (-not (Test-Path -LiteralPath (Join-Path $payload.Directory $payload.Executable) -PathType Leaf)) {
        throw "The single-file executable is missing: $($payload.Executable)"
    }
    $externalCode = @(Get-ChildItem -LiteralPath $payload.Directory -Recurse -File -Force | Where-Object {
        $_.Extension -eq '.dll' -or ($_.Extension -eq '.exe' -and $_.Name -ne $payload.Executable) -or
        $_.Name -like '*.deps.json' -or $_.Name -like '*.runtimeconfig*.json'
    })
    if ($externalCode.Count -ne 0) { throw "Executable payload escaped the frozen single-file hash: $($externalCode.FullName -join ', ')" }
}

$smokeProject = Join-Path $repositoryRoot 'tools\ExcelDiffTracker.Smoke\ExcelDiffTracker.Smoke.csproj'
dotnet publish $smokeProject -c $Configuration -r win-arm64 --self-contained true --no-restore -o $smokeTool
if ($LASTEXITCODE -ne 0) { throw 'ARM64 smoke-tool publish failed.' }

Copy-Item (Join-Path $repositoryRoot 'LICENSE') $publish
Copy-Item (Join-Path $repositoryRoot 'THIRD-PARTY-NOTICES.md') $publish
$licenseDirectory = Join-Path $publish 'licenses'
$dotnetLicenseDirectory = Join-Path $licenseDirectory 'dotnet'
New-Item -ItemType Directory -Path $dotnetLicenseDirectory -Force | Out-Null
Copy-Item (Join-Path $repositoryRoot 'licenses\Apache-2.0.txt') (Join-Path $licenseDirectory 'SQLitePCLRaw-2.1.12-Apache-2.0.txt')

# Single-file publishing embeds runtime/dependency JSON. Resolve the matching
# SDK build metadata from this exact configuration/RID for the bundled licenses.
$metadataJson = & dotnet msbuild $appProject "-p:Configuration=$Configuration" '-p:RuntimeIdentifier=win-arm64' '-getProperty:ProjectRuntimeConfigFilePath,ProjectDepsFilePath'
if ($LASTEXITCODE -ne 0) { throw 'Could not resolve the published application metadata paths.' }
$metadata = ($metadataJson | Out-String | ConvertFrom-Json).Properties
$runtimeConfig = Get-Content -LiteralPath $metadata.ProjectRuntimeConfigFilePath -Raw | ConvertFrom-Json
$coreFramework = @($runtimeConfig.runtimeOptions.includedFrameworks | Where-Object name -eq 'Microsoft.NETCore.App') | Select-Object -First 1
$desktopFramework = @($runtimeConfig.runtimeOptions.includedFrameworks | Where-Object name -eq 'Microsoft.WindowsDesktop.App') | Select-Object -First 1
if (-not $coreFramework -or -not $desktopFramework) { throw 'Published runtime versions could not be determined.' }

$nugetRoot = if ($env:NUGET_PACKAGES) { $env:NUGET_PACKAGES } else { Join-Path $env:USERPROFILE '.nuget\packages' }
$coreRuntimePackage = Join-Path $nugetRoot "microsoft.netcore.app.runtime.win-arm64\$($coreFramework.version)"
$desktopRuntimePackage = Join-Path $nugetRoot "microsoft.windowsdesktop.app.runtime.win-arm64\$($desktopFramework.version)"
$requiredRuntimeNotices = @(
    @{ Source = Join-Path $coreRuntimePackage 'LICENSE.TXT'; Destination = "Microsoft.NETCore.App-$($coreFramework.version)-LICENSE.txt" },
    @{ Source = Join-Path $coreRuntimePackage 'THIRD-PARTY-NOTICES.TXT'; Destination = "Microsoft.NETCore.App-$($coreFramework.version)-THIRD-PARTY-NOTICES.txt" },
    @{ Source = Join-Path $desktopRuntimePackage 'LICENSE'; Destination = "Microsoft.WindowsDesktop.App-$($desktopFramework.version)-LICENSE.txt" }
)
foreach ($notice in $requiredRuntimeNotices) {
    if (-not (Test-Path $notice.Source)) { throw "Required runtime notice was not found: $($notice.Source)" }
    Copy-Item $notice.Source (Join-Path $dotnetLicenseDirectory $notice.Destination)
}

$dependencyManifest = Get-Content -LiteralPath $metadata.ProjectDepsFilePath -Raw | ConvertFrom-Json
$windowsSdkLibrary = @($dependencyManifest.libraries.PSObject.Properties.Name | Where-Object {
    $_ -match '^(runtimepack\.)?Microsoft\.Windows\.SDK\.NET\.Ref/'
}) | Select-Object -First 1
if (-not $windowsSdkLibrary) { throw 'The published Windows SDK runtime pack could not be identified.' }
$windowsSdkVersion = $windowsSdkLibrary.Split('/')[-1]
$windowsSdkLicense = Join-Path $repositoryRoot "licenses\Microsoft-Windows-SDK-$windowsSdkVersion-License.rtf"
if (-not (Test-Path $windowsSdkLicense)) {
    throw "The exact Windows SDK $windowsSdkVersion license is not checked in: $windowsSdkLicense"
}
$expectedWindowsSdkLicenseHashes = @{
    '10.0.26100.57' = 'DD07EB178E00C6BBA4148457FC00FF77CD4887EB521D504186FE59C9EC8BBE62'
}
$expectedWindowsSdkLicenseHash = $expectedWindowsSdkLicenseHashes[$windowsSdkVersion]
if (-not $expectedWindowsSdkLicenseHash) {
    throw "No approved license hash is recorded for Windows SDK runtime pack $windowsSdkVersion."
}
$actualWindowsSdkLicenseHash = (Get-FileHash $windowsSdkLicense -Algorithm SHA256).Hash.ToUpperInvariant()
if ($actualWindowsSdkLicenseHash -ne $expectedWindowsSdkLicenseHash) {
    throw "The Windows SDK $windowsSdkVersion license hash does not match the approved source."
}
Copy-Item $windowsSdkLicense (Join-Path $dotnetLicenseDirectory "Microsoft-Windows-SDK-$windowsSdkVersion-License.rtf")

if (-not $SkipInstaller) {
    $candidateCompilers = @(
        (Join-Path $env:LOCALAPPDATA 'Programs\Inno Setup 7\ISCC.exe'),
        (Join-Path $env:ProgramFiles 'Inno Setup 7\ISCC.exe'),
        (Join-Path ${env:ProgramFiles(x86)} 'Inno Setup 7\ISCC.exe')
    ) | Where-Object { $_ -and (Test-Path $_) }
    $compiler = $candidateCompilers | Select-Object -First 1
    if (-not $compiler) {
        throw 'Inno Setup 7 was not found. Install it or rerun with -SkipInstaller.'
    }

    $installerScript = Join-Path $repositoryRoot 'installer\ExcelDiffTracker.iss'
    & $compiler "/DMyAppVersion=$Version" $installerScript
    if ($LASTEXITCODE -ne 0) { throw 'Installer compilation failed.' }

    $installer = Join-Path $artifacts 'installer\ExcelDiffTracker-Setup-arm64.exe'
    if (-not (Test-Path $installer)) { throw 'The expected installer was not produced.' }
    Copy-Item $installer $release

    $installerHash = (Get-FileHash $installer -Algorithm SHA256).Hash.ToUpperInvariant()
    if (Test-Path $wingetManifest) {
        $manifestText = [System.IO.File]::ReadAllText($wingetManifest)
        $updatedManifest = [System.Text.RegularExpressions.Regex]::Replace(
            $manifestText,
            '(?m)^(\s*InstallerSha256:\s*)[A-Fa-f0-9]{64}\s*$',
            "`${1}$installerHash")
        if ($updatedManifest -eq $manifestText -and $manifestText -notmatch [System.Text.RegularExpressions.Regex]::Escape($installerHash)) {
            throw 'The WinGet installer hash field could not be updated.'
        }
        [System.IO.File]::WriteAllText($wingetManifest, $updatedManifest, [System.Text.UTF8Encoding]::new($false))
    }
}

Copy-Item (Join-Path $repositoryRoot 'LICENSE') $release
Copy-Item (Join-Path $repositoryRoot 'THIRD-PARTY-NOTICES.md') $release
$releaseNotes = Join-Path $repositoryRoot "docs\RELEASE_NOTES_$Version.md"
if (-not (Test-Path $releaseNotes)) { throw "Release notes were not found: $releaseNotes" }
Copy-Item $releaseNotes $release
Compress-Archive -Path (Join-Path $licenseDirectory '*') -DestinationPath (Join-Path $release 'THIRD-PARTY-LICENSES.zip') -CompressionLevel Optimal

$releaseFiles = Get-ChildItem $release -File | Where-Object { $_.Name -ne 'SHA256SUMS.txt' } | Sort-Object Name
$checksumLines = foreach ($file in $releaseFiles) {
    $hash = (Get-FileHash $file.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
    "$hash  $($file.Name)"
}
$checksumText = ($checksumLines -join "`n") + "`n"
[System.IO.File]::WriteAllText(
    (Join-Path $release 'SHA256SUMS.txt'),
    $checksumText,
    [System.Text.UTF8Encoding]::new($false))

Write-Host "Release artifacts: $release"
