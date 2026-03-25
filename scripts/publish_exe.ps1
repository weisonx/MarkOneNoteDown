param(
    [string]$Configuration = "Release",
    [string]$Runtime = "win-x64",
    [string]$OutputDir = "artifacts\\exe",
    [switch]$CleanOutput = $true,
    [switch]$SelfContained = $true
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
}
$resolvedOutput = (Resolve-Path $OutputDir).Path

if ($CleanOutput -and (Test-Path $resolvedOutput)) {
    Write-Host "Cleaning output directory: $resolvedOutput"
    Get-ChildItem -Path $resolvedOutput -Force | Remove-Item -Force -Recurse
}

$project = "src\\MarkOneNoteDown.App\\MarkOneNoteDown.App.csproj"

$sc = if ($SelfContained) { "true" } else { "false" }

dotnet publish $project -c $Configuration -r $Runtime --self-contained $sc `
    -p:WindowsPackageType=None `
    -p:PublishSingleFile=false `
    -p:IncludeNativeLibrariesForSelfExtract=true `
    -p:AppxPackageSigningEnabled=false `
    -p:PublishReadyToRun=false `
    -o $resolvedOutput

Write-Host "EXE output directory: $resolvedOutput"
