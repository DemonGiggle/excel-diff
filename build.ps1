[CmdletBinding()]
param(
    [ValidateSet("Release", "Debug")]
    [string]$Configuration = "Release",

    [switch]$SkipTests,

    [ValidatePattern('^\d+\.\d+\.\d+$')]
    [string]$Version,

    [switch]$LockedMode
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$repositoryRoot = $PSScriptRoot
$solutionPath = Join-Path $repositoryRoot "ExcelDiff.sln"
$publishScript = Join-Path $repositoryRoot "scripts\publish.ps1"
$localDotnet = Join-Path $repositoryRoot ".tools\dotnet-complete\dotnet.exe"
$localCliHome = Join-Path $repositoryRoot ".tools\cli-home"

function Find-DotNet {
    if (Test-Path -LiteralPath $localDotnet -PathType Leaf) {
        New-Item -ItemType Directory -Path $localCliHome -Force | Out-Null
        $env:DOTNET_CLI_HOME = $localCliHome
        return $localDotnet
    }

    $command = Get-Command dotnet -ErrorAction SilentlyContinue
    if ($null -ne $command) {
        return $command.Source
    }

    throw @"
The .NET 8 SDK was not found.

Install it with:
  winget install --id Microsoft.DotNet.SDK.8 --source winget

Then close and reopen PowerShell and run:
  .\build.ps1

The SDK download is also available at https://dotnet.microsoft.com/download/dotnet/8.0
"@
}

$dotnet = Find-DotNet
$env:DOTNET_SKIP_FIRST_TIME_EXPERIENCE = "1"
$env:DOTNET_CLI_TELEMETRY_OPTOUT = "1"

Write-Host "Using .NET SDK: $dotnet" -ForegroundColor Cyan
& $dotnet --version

Write-Host "`nRestoring packages..." -ForegroundColor Cyan
$restoreArguments = @("restore", $solutionPath, "--configfile", (Join-Path $repositoryRoot "NuGet.Config"))
if ($LockedMode) {
    $restoreArguments += "--locked-mode"
}
& $dotnet @restoreArguments
if ($LASTEXITCODE -ne 0) { throw "Package restore failed with exit code $LASTEXITCODE." }

if (-not $SkipTests) {
    Write-Host "`nRunning tests..." -ForegroundColor Cyan
    & $dotnet test $solutionPath --configuration $Configuration --no-restore
    if ($LASTEXITCODE -ne 0) { throw "Tests failed with exit code $LASTEXITCODE." }
}

Write-Host "`nBuilding the portable Windows application..." -ForegroundColor Cyan
$publishArguments = @{
    Configuration = $Configuration
    DotNetPath = $dotnet
    LockedMode = $LockedMode
}
if ($Version) {
    $publishArguments.Version = $Version
}
& $publishScript @publishArguments
if ($LASTEXITCODE -ne 0) { throw "Publishing failed with exit code $LASTEXITCODE." }

Write-Host "`nBuild completed successfully." -ForegroundColor Green
Write-Host "Application: artifacts\ExcelDiff-win-x64\ExcelDiff.exe"
Write-Host "Portable ZIP: artifacts\ExcelDiff-win-x64-portable.zip"
