param(
    [string]$MsixPath = ""
)

$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($MsixPath)) {
    $msix = Get-ChildItem -Recurse -Path "artifacts\\msix" -Filter *.msix -File |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
    if ($null -eq $msix) {
        throw "No MSIX found under artifacts\\msix. Run scripts\\package_msix.ps1 first."
    }
    $MsixPath = $msix.FullName
}

Write-Host "Installing MSIX: $MsixPath"
Add-AppxPackage -Path $MsixPath
