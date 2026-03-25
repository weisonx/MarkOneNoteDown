param(
    [string]$Configuration = "Release",
    [string]$OutputDir = "artifacts\\msix",
    [string]$CertPath = "",
    [string]$CertPassword = "",
    [switch]$TrustCertificate = $true,
    [switch]$CleanOutput = $true,
    [switch]$InstallAfter = $false
)

$ErrorActionPreference = "Stop"

function Convert-ToPlainText {
    param([System.Security.SecureString]$Secure)
    $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Secure)
    try {
        return [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
    }
    finally {
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    }
}

if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
}
$resolvedOutput = (Resolve-Path $OutputDir).Path

if ($CleanOutput -and (Test-Path $resolvedOutput)) {
    Write-Host "Cleaning output directory: $resolvedOutput"
    Get-ChildItem -Path $resolvedOutput -Force | Remove-Item -Force -Recurse
}

if ([string]::IsNullOrWhiteSpace($CertPath)) {
    $certDir = "certs"
    if (-not (Test-Path $certDir)) {
        New-Item -ItemType Directory -Force -Path $certDir | Out-Null
    }
    $CertPath = Join-Path $certDir "MarkOneNoteDown.pfx"
}

$certDirPath = Split-Path -Parent $CertPath
if (-not (Test-Path $certDirPath)) {
    New-Item -ItemType Directory -Force -Path $certDirPath | Out-Null
}

$resolvedCertPath = (Resolve-Path $certDirPath).Path + "\\" + (Split-Path -Leaf $CertPath)
$cerPath = [System.IO.Path]::ChangeExtension($resolvedCertPath, ".cer")

if (-not (Test-Path $resolvedCertPath)) {
    $securePwd = if ([string]::IsNullOrWhiteSpace($CertPassword)) {
        Read-Host "Enter certificate password" -AsSecureString
    }
    else {
        ConvertTo-SecureString $CertPassword -AsPlainText -Force
    }

    $cert = New-SelfSignedCertificate -Type CodeSigning -Subject "CN=User Name" -CertStoreLocation "Cert:\\CurrentUser\\My"
    Export-PfxCertificate -Cert $cert -FilePath $resolvedCertPath -Password $securePwd | Out-Null
    Export-Certificate -Cert $cert -FilePath $cerPath | Out-Null

    Write-Host "Created signing certificate:"
    Write-Host "  PFX: $resolvedCertPath"
    Write-Host "  CER: $cerPath"

    $CertPassword = Convert-ToPlainText $securePwd
}
elseif ([string]::IsNullOrWhiteSpace($CertPassword)) {
    $securePwd = Read-Host "Enter certificate password" -AsSecureString
    $CertPassword = Convert-ToPlainText $securePwd
}

if ($TrustCertificate -and (Test-Path $cerPath)) {
    Write-Host "Trusting certificate in CurrentUser\\TrustedPeople..."
    Import-Certificate -FilePath $cerPath -CertStoreLocation Cert:\CurrentUser\TrustedPeople | Out-Null
    Write-Host "Trusting certificate in CurrentUser\\Root..."
    Import-Certificate -FilePath $cerPath -CertStoreLocation Cert:\CurrentUser\Root | Out-Null
}

$project = "src\\MarkOneNoteDown.App\\MarkOneNoteDown.App.csproj"

dotnet publish $project -c $Configuration -p:Platform=x64 -p:PublishProfile=win-x64.pubxml `
    -p:WindowsPackageType=MSIX `
    -p:GenerateAppxPackageOnBuild=true `
    -p:UapAppxPackageBuildMode=SideloadOnly `
    -p:AppxBundle=Never `
    -p:AppxPackageDir="$resolvedOutput\\" `
    -p:AppxPackageSigningEnabled=true `
    -p:PackageCertificateKeyFile="$resolvedCertPath" `
    -p:PackageCertificatePassword="$CertPassword"

Write-Host "MSIX output directory: $resolvedOutput"

$msix = Get-ChildItem -Recurse -Path $resolvedOutput -Filter *.msix -File |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1
if ($null -eq $msix) {
    throw "No MSIX found in $resolvedOutput."
}

$sig = Get-AuthenticodeSignature $msix.FullName
if ($sig.Status -ne "Valid") {
    throw "MSIX is not signed correctly. Status: $($sig.Status). Message: $($sig.StatusMessage)"
}

if ($InstallAfter) {
    Write-Host "Installing MSIX: $($msix.FullName)"
    Add-AppxPackage -Path $msix.FullName
}
