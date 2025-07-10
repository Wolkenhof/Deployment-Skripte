<#
############################## Wolkenhof ##############################
Purpose : SwyxIt! Client Installation
Created : 07.07.2025
Updated : 10.07.2025
Version : 1.1
Source  : https://github.com/Wolkenhof/Deployment-Skripte
Author  : jgu
Company : Wolkenhof GmbH
############################## Wolkenhof ##############################
#>

<# Settings #>

$swyxVersion = '14.21.0.0'                          # SwyxIt! Client Version
$serverIP = '10.10.1.103'                           # SwyxIt! Server IP address
#$pbxUser = ''                                       # optional; leave blank if not needed 
$useTrustedAuthentication = 0                       # Set this to 1 if using Windows Authentification
$remoteConnectorAuth = 'swyx.example.com:9101'     # Remote Connection Authentification (PublicAuthServerName)
$remoteConnectorServer = 'swyx.example.com:16203'  # Remote Connection Server (PublicAuthServerName)

# !! Advanced Settings - Change only if you know what you're doing !!
$zipFileUrl = "https://downloads.enreach.de/download/swyxit!_${swyxVersion}_64bit_german.zip"
$zipOutput = (Join-Path $env:TEMP "SwyxDeployment")
$msiFileName = 'SwyxIt!German64.msi'

<# Settings #>

function Get-ZipAndExtract {
param(
        [Parameter(Mandatory=$true)]
        [string]$ZipUrl,

        [Parameter(Mandatory=$true)]
        [string]$DestinationPath
    )

    $tempZipFile = Join-Path $env:TEMP ([System.IO.Path]::GetRandomFileName() + ".zip")

    Write-Host "Starting download of '$ZipUrl' to '$tempZipFile'..."
    try {
		$wc = New-Object net.webclient
		$wc.Downloadfile($ZipUrl, $tempZipFile)
        Write-Host "Download complete."
    } catch {
        Write-Error "Error downloading the file: $($_.Exception.Message)"
        return
    }

    if (-not (Test-Path $DestinationPath)) {
        New-Item -ItemType Directory -Path $DestinationPath -Force | Out-Null
        Write-Host "Destination folder '$DestinationPath' created."
    } else {
        Write-Host "Destination folder '$DestinationPath' already exists."
    }

    Write-Host "Extracting '$tempZipFile' to '$DestinationPath'..."
    try {
        Expand-Archive -Path $tempZipFile -DestinationPath $DestinationPath -Force
        Write-Host "Extraction successful!"
    } catch {
        Write-Error "Error extracting the file: $($_.Exception.Message)"
        return
    }

    Write-Host "Deleting temporary ZIP file '$tempZipFile'..."
    try {
        Remove-Item -Path $tempZipFile -Force -ErrorAction Stop
        Write-Host "Temporary ZIP file successfully deleted."
    } catch {
        Write-Warning "Error deleting the temporary ZIP file: $($_.Exception.Message)"
    }

    Write-Host "Content extracted to: $DestinationPath"
    return $DestinationPath
}

function Install-SwyxItAndSetRegistry {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ExtractedFolderPath
    )

    $msiPath = Join-Path $ExtractedFolderPath $msiFileName
    Write-Host "Searching for MSI file at: $msiPath"

    if (Test-Path $msiPath) {
        Write-Host "Starting MSI installation for '$msiPath'..."
        try {
            $process = Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$msiPath`" /qn /norestart" -Wait -PassThru
            if ($process.ExitCode -eq 0) {
                Write-Host "MSI installation completed successfully."
            } else {
                Write-Error "MSI installation failed with exit code: $($process.ExitCode)"
                return
            }
        } catch {
            Write-Error "An error occurred during MSI installation: $($_.Exception.Message)"
            return
        }
    } else {
        Write-Error "MSI file not found at '$msiPath'. Skipping installation."
        return
    }

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
}

try {
    $extractedFolder = Get-ZipAndExtract -ZipUrl $zipFileUrl -DestinationPath $zipOutput 
    if ($extractedFolder) {
        Write-Host "`Download and Extraction completed. Extracted content is in: $extractedFolder"
        Install-SwyxItAndSetRegistry -ExtractedFolderPath $extractedFolder
    } else {
        Write-Error "Download and Extraction failed. Aborting further steps."
    }
} catch {
    Write-Error "An unexpected error occurred: $($_.Exception.Message)"
}
