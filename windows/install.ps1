# Sets up the Claude AI add-in for Excel on Windows without npm install:
#  1. creates and trusts a localhost HTTPS certificate (current user only),
#  2. registers the add-in manifest with Office (sideload),
#  3. leaves server startup to install-windows.cmd.
$ErrorActionPreference = "Stop"

$here      = Split-Path -Parent $MyInvocation.MyCommand.Path
$manifest  = Join-Path $here "addin\manifest.xml"
$dataDir   = Join-Path $env:LOCALAPPDATA "ClaudeExcelAddin"
$pfxPath   = Join-Path $dataDir "localhost.pfx"
$cerPath   = Join-Path $dataDir "localhost.cer"
$friendly  = "Claude Excel Add-in (localhost)"
# Protects only the local copy of a localhost-only certificate.
$pfxPass   = "claude-excel-localhost"
$addinId   = "6f2c8a8e-3a5b-4f63-9a0e-2d7c1b9e4a51"
$devKey    = "HKCU:\Software\Microsoft\Office\16.0\Wef\Developer"

New-Item -ItemType Directory -Force -Path $dataDir | Out-Null

# --- 1. Certificate -------------------------------------------------------
$cert = Get-ChildItem Cert:\CurrentUser\My |
  Where-Object { $_.FriendlyName -eq $friendly -and $_.NotAfter -gt (Get-Date).AddDays(30) } |
  Select-Object -First 1
if (-not $cert) {
  Write-Host "Creating a localhost certificate..."
  $cert = New-SelfSignedCertificate -DnsName "localhost" -CertStoreLocation "Cert:\CurrentUser\My" `
    -FriendlyName $friendly -NotAfter (Get-Date).AddYears(5) -KeyExportPolicy Exportable `
    -KeyAlgorithm RSA -KeyLength 2048 -HashAlgorithm SHA256
}

$trusted = Get-ChildItem Cert:\CurrentUser\Root | Where-Object { $_.Thumbprint -eq $cert.Thumbprint }
if (-not $trusted) {
  Write-Host ""
  Write-Host "Windows will now ask whether to install a certificate for 'localhost'." -ForegroundColor Yellow
  Write-Host "Click YES. It only lets Excel load the add-in from this computer." -ForegroundColor Yellow
  Export-Certificate -Cert $cert -FilePath $cerPath | Out-Null
  Import-Certificate -FilePath $cerPath -CertStoreLocation Cert:\CurrentUser\Root | Out-Null
}

$secure = ConvertTo-SecureString -String $pfxPass -Force -AsPlainText
try {
  # AES-256 keeps the file readable by current Node.js (OpenSSL 3).
  Export-PfxCertificate -Cert $cert -FilePath $pfxPath -Password $secure -CryptoAlgorithmOption AES256_SHA256 | Out-Null
} catch {
  Export-PfxCertificate -Cert $cert -FilePath $pfxPath -Password $secure | Out-Null
}

# --- 2. Sideload into Office ------------------------------------------------
if (-not (Test-Path $devKey)) { New-Item -Path $devKey -Force | Out-Null }
New-ItemProperty -Path $devKey -Name $addinId -Value $manifest -PropertyType String -Force | Out-Null
Write-Host "Add-in registered with Excel." -ForegroundColor Green
