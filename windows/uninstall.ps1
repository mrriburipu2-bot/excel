# Removes the sideloaded add-in and its localhost certificate.
$addinId  = "6f2c8a8e-3a5b-4f63-9a0e-2d7c1b9e4a51"
$devKey   = "HKCU:\Software\Microsoft\Office\16.0\Wef\Developer"
$friendly = "Claude Excel Add-in (localhost)"

Remove-ItemProperty -Path $devKey -Name $addinId -ErrorAction SilentlyContinue
$certs = Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.FriendlyName -eq $friendly }
foreach ($c in $certs) {
  Get-ChildItem Cert:\CurrentUser\Root | Where-Object { $_.Thumbprint -eq $c.Thumbprint } | Remove-Item -ErrorAction SilentlyContinue
  Remove-Item $c.PSPath -ErrorAction SilentlyContinue
}
Remove-Item (Join-Path $env:LOCALAPPDATA "ClaudeExcelAddin") -Recurse -Force -ErrorAction SilentlyContinue
Write-Host "Claude AI add-in removed. Restart Excel to finish."
