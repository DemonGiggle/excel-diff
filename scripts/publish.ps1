param(
    [string]$Configuration = "Release",
    [string]$DotNetPath = "dotnet"
)

$ErrorActionPreference = "Stop"
$repositoryRoot = Split-Path -Parent $PSScriptRoot
$projectPath = Join-Path $repositoryRoot "src\ExcelDiff.App\ExcelDiff.App.csproj"
$artifactRoot = Join-Path $repositoryRoot "artifacts"
$publishDirectory = Join-Path $artifactRoot "ExcelDiff-win-x64"
$zipPath = Join-Path $artifactRoot "ExcelDiff-win-x64-portable.zip"

New-Item -ItemType Directory -Path $artifactRoot -Force | Out-Null
& $DotNetPath restore $projectPath -r win-x64 --configfile (Join-Path $repositoryRoot "NuGet.Config")
if ($LASTEXITCODE -ne 0) { throw "Runtime package restore failed with exit code $LASTEXITCODE." }
& $DotNetPath publish $projectPath -c $Configuration -r win-x64 --self-contained true --no-restore -p:PublishSingleFile=false -p:DebugType=None -p:DebugSymbols=false -o $publishDirectory
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
