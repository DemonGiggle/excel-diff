param(
    [string]$Configuration = "Release"
)

$ErrorActionPreference = "Stop"
$repositoryRoot = Split-Path -Parent $PSScriptRoot
$projectPath = Join-Path $repositoryRoot "src\ExcelDiff.App\ExcelDiff.App.csproj"
$artifactRoot = Join-Path $repositoryRoot "artifacts"
$publishDirectory = Join-Path $artifactRoot "ExcelDiff-win-x64"
$zipPath = Join-Path $artifactRoot "ExcelDiff-win-x64-portable.zip"

New-Item -ItemType Directory -Path $artifactRoot -Force | Out-Null
dotnet restore $projectPath -r win-x64 --configfile (Join-Path $repositoryRoot "NuGet.Config")
dotnet publish $projectPath -c $Configuration -r win-x64 --self-contained true --no-restore -p:PublishSingleFile=false -p:DebugType=None -p:DebugSymbols=false -o $publishDirectory

if (Test-Path -LiteralPath $zipPath) {
    Remove-Item -LiteralPath $zipPath -Force
}
Compress-Archive -Path (Join-Path $publishDirectory "*") -DestinationPath $zipPath -CompressionLevel Optimal
Write-Output "Portable app: $publishDirectory"
Write-Output "Portable ZIP: $zipPath"
