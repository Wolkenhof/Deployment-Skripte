param (
    [string]$ScriptPath
)

Write-Host "Pfad: $ScriptPath"

if (Test-Path -Path $ScriptPath) {
    Write-Host "Starte Signatur ..."
    
    $timeStampServer = "http://time.certum.pl"
    $CertificateThumbprint = "e302d41d12ff036dbd37fa03d5c85f376edf659b"
    # RFC 2253: CN=Jonas G\C3\BCnner,O=Jonas G\C3\BCnner,L=Uelzen,ST=Niedersachsen,C=DE

    # Get the Certificate from Cert Store
    $CodeSignCert = Get-ChildItem -Path Cert:\CurrentUser\My | Where-Object {$_.Thumbprint -eq $CertificateThumbprint}
    
    # Sign the PS1 file
    Set-AuthenticodeSignature -FilePath $ScriptPath -Certificate $CodeSignCert -TimestampServer $timeStampServer -HashAlgorithm sha256 
}
else {
    Write-Host "Fehler"
}

