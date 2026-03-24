param(
    [string]$Configuration = "Release",
    [string]$OutputDir = "artifacts\\msix",
    [string]$CertPath = "",
    [string]$CertPassword = "",
    [switch]$TrustCertificate = $true
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

if ([string]::IsNullOrWhiteSpace($CertPath)) {
    $certDir = "certs"
    if (-not (Test-Path $certDir)) {
        New-Item -ItemType Directory -Force -Path $certDir | Out-Null
    }
    $CertPath = Join-Path $certDir "MarkOneNoteDown.pfx"
}

$cerPath = [System.IO.Path]::ChangeExtension($CertPath, ".cer")

if (-not (Test-Path $CertPath)) {
    $securePwd = if ([string]::IsNullOrWhiteSpace($CertPassword)) {
        Read-Host "Enter certificate password" -AsSecureString
    }
    else {
        ConvertTo-SecureString $CertPassword -AsPlainText -Force
    }

    $cert = New-SelfSignedCertificate -Type CodeSigning -Subject "CN=User Name" -CertStoreLocation "Cert:\\CurrentUser\\My"
    Export-PfxCertificate -Cert $cert -FilePath $CertPath -Password $securePwd | Out-Null
    Export-Certificate -Cert $cert -FilePath $cerPath | Out-Null

    Write-Host "Created signing certificate:"
    Write-Host "  PFX: $CertPath"
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
}

$project = "src\\MarkOneNoteDown.App\\MarkOneNoteDown.App.csproj"

dotnet publish $project -c $Configuration -p:Platform=x64 -p:PublishProfile=win-x64.pubxml `
    -p:WindowsPackageType=MSIX `
    -p:GenerateAppxPackageOnBuild=true `
    -p:UapAppxPackageBuildMode=SideloadOnly `
    -p:AppxBundle=Never `
    -p:AppxPackageDir="$resolvedOutput\\" `
    -p:PackageCertificateKeyFile="$CertPath" `
    -p:PackageCertificatePassword="$CertPassword"

Write-Host "MSIX output directory: $resolvedOutput"
