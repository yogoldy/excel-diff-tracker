[CmdletBinding()]
param([Parameter(Mandatory)] [string] $ExecutablePath)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$executable = Get-Item -LiteralPath $ExecutablePath -ErrorAction Stop
if ($executable.PSIsContainer -or $executable.Extension -ne '.exe') { throw 'A published executable is required.' }
$looseRuntime = @(Get-ChildItem -LiteralPath $executable.DirectoryName -Recurse -File -Force -ErrorAction Stop | Where-Object {
    $_.Extension -eq '.dll' -or $_.Name -like '*.deps.json' -or $_.Name -like '*.runtimeconfig*.json'
})
if ($looseRuntime.Count -ne 0) {
    throw "Loose executable dependencies are outside the frozen single-file hash: $($looseRuntime.FullName -join ', ')"
}
Write-Output "SINGLE_FILE_PAYLOAD_VALID|$($executable.FullName)"
