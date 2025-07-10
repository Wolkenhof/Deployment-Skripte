$serverIP = '10.10.1.103'                           # SwyxIt! Server IP address
#$pbxUser = ''                                       # optional; leave blank if not needed 
$useTrustedAuthentication = 0                       # Set this to 1 if using Windows Authentification
$remoteConnectorAuth = 'swyx.example.com:9101'     # Remote Connection Authentification (PublicAuthServerName)
$remoteConnectorServer = 'swyx.example.com:16203'  # Remote Connection Server (PublicAuthServerName)

#################################################################################

$regPath = "HKCU:\SOFTWARE\Swyx\SwyxIt!\CurrentVersion\Options"
Write-Host "Setting Registry Keys under '$regPath'..."

try {
    if (-not (Test-Path $regPath)) {
        New-Item -Path $regPath -Force | Out-Null
        Write-Host "Registry path '$regPath' created."
    }

    Set-ItemProperty -Path $regPath -Name "PbxServer" -Value $serverIP -Force
    #Set-ItemProperty -Path $regPath -Name "PbxUser" -Value $pbxUser -Force
    Set-ItemProperty -Path $regPath -Name "TrustedAuthentication" -Value $useTrustedAuthentication -Type DWord -Force
    Set-ItemProperty -Path $regPath -Name "PublicAuthServerName" -Value $remoteConnectorAuth -Type String -Force
    Set-ItemProperty -Path $regPath -Name "PublicServerName" -Value $remoteConnectorServer -Type String -Force

    Write-Host "Registry keys set successfully."
} catch {
    Write-Error "Error setting registry keys: $($_.Exception.Message)"
    return
}
