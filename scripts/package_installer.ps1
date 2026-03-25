param(
    [string]$Configuration = "Release",
    [string]$Runtime = "win-x64",
    [string]$OutputDir = "artifacts\\installer",
    [string]$PublishDir = "artifacts\\exe",
    [switch]$CleanOutput = $true,
    [switch]$SelfContained = $true
)

$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path ".").Path
$publishScript = Join-Path $repoRoot "scripts\\publish_exe.ps1"

if (-not (Test-Path $publishScript)) {
    throw "Missing script: $publishScript"
}

& $publishScript -Configuration $Configuration -Runtime $Runtime -OutputDir $PublishDir -SelfContained:$SelfContained

if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
}
$resolvedOutput = (Resolve-Path $OutputDir).Path

if ($CleanOutput -and (Test-Path $resolvedOutput)) {
    Write-Host "Cleaning output directory: $resolvedOutput"
    Get-ChildItem -Path $resolvedOutput -Force | Remove-Item -Force -Recurse
}

$manifestPath = Join-Path $repoRoot "src\\MarkOneNoteDown.App\\Package.appxmanifest"
if (-not (Test-Path $manifestPath)) {
    throw "Package manifest not found: $manifestPath"
}

$manifest = [xml](Get-Content $manifestPath)
$version = $manifest.Package.Identity.Version

$iscc = Get-Command "ISCC.exe" -ErrorAction SilentlyContinue
if ($null -eq $iscc) {
    Write-Host "ISCC.exe not found. Please install Inno Setup and make sure ISCC.exe is in PATH."
    Write-Host "Publish output is ready in: $PublishDir"
    return
}

$issPath = Join-Path $repoRoot "scripts\\installer_exe.iss"
if (-not (Test-Path $issPath)) {
    throw "Installer script not found: $issPath"
}

$resolvedPublish = (Resolve-Path $PublishDir).Path

& $iscc.Path "/DAppVersion=$version" "/DSourceDir=$resolvedPublish" "/DOutputDir=$resolvedOutput" $issPath

Write-Host "Installer output directory: $resolvedOutput"
