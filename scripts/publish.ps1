param(
    [string]$Configuration = "Release",
    [string]$DotNetPath = "dotnet",
    [ValidatePattern('^\d+\.\d+\.\d+$')]
    [string]$Version,
    [switch]$LockedMode
)

$ErrorActionPreference = "Stop"
$repositoryRoot = Split-Path -Parent $PSScriptRoot
$projectPath = Join-Path $repositoryRoot "src\ExcelDiff.App\ExcelDiff.App.csproj"
$artifactRoot = Join-Path $repositoryRoot "artifacts"
$publishDirectory = Join-Path $artifactRoot "ExcelDiff-win-x64"
$zipPath = Join-Path $artifactRoot "ExcelDiff-win-x64-portable.zip"

New-Item -ItemType Directory -Path $artifactRoot -Force | Out-Null
$restoreArguments = @("restore", $projectPath, "-r", "win-x64", "--configfile", (Join-Path $repositoryRoot "NuGet.Config"))
if ($LockedMode) {
    $restoreArguments += "--locked-mode"
}
& $DotNetPath @restoreArguments
if ($LASTEXITCODE -ne 0) { throw "Runtime package restore failed with exit code $LASTEXITCODE." }

if (Test-Path -LiteralPath $publishDirectory) {
    Remove-Item -LiteralPath $publishDirectory -Recurse -Force
}

$publishArguments = @(
    "publish", $projectPath,
    "-c", $Configuration,
    "-r", "win-x64",
    "--self-contained", "true",
    "--no-restore",
    "-p:PublishSingleFile=false",
    "-p:DebugType=None",
    "-p:DebugSymbols=false",
    "-o", $publishDirectory
)
if ($Version) {
    $publishArguments += @(
        "-p:Version=$Version",
        "-p:AssemblyVersion=$Version.0",
        "-p:FileVersion=$Version.0",
        "-p:InformationalVersion=$Version"
    )
}
& $DotNetPath @publishArguments
if ($LASTEXITCODE -ne 0) { throw "Publishing failed with exit code $LASTEXITCODE." }

if (Test-Path -LiteralPath $zipPath) {
    Remove-Item -LiteralPath $zipPath -Force
}
try {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::CreateFromDirectory(
        $publishDirectory,
        $zipPath,
        [System.IO.Compression.CompressionLevel]::Optimal,
        $false
    )
}
catch {
    if (Test-Path -LiteralPath $zipPath) {
        Remove-Item -LiteralPath $zipPath -Force
    }
    throw "The portable ZIP could not be created. Close ExcelDiff.exe if it is running and try again. $($_.Exception.Message)"
}
Write-Output "Portable app: $publishDirectory"
Write-Output "Portable ZIP: $zipPath"
