<#
############################## Wolkenhof ##############################
Purpose : Install or Update OneDrive on Windows Server
Created : 27.11.2024
Source  : https://github.com/Wolkenhof/Deployment-Skripte
Author  : jgu
Company : Wolkenhof GmbH
############################## Wolkenhof ##############################
#>
Add-Type -AssemblyName System.Net.Http

function DownloadOneDrive($DownloadUrl, $FileName) {
    function Show-Progress {
        param (
            [int]$PercentComplete
        )
        Write-Progress -Activity "Downloading OneDrive" -Status "$PercentComplete% Complete" -PercentComplete $PercentComplete
    }

    Write-Output "[i] Downloading OneDrive ..."
    try {
        $destinationPath = "C:\Wolkenhof\Source\OneDrive\$FileName"
        if ($IsClassic) {
            $destinationPath = "C:\Wolkenhof\Source\OneDrive\$FileName"
        }

        $destinationDir = [System.IO.Path]::GetDirectoryName($destinationPath)
        if (-not (Test-Path -Path $destinationDir)) {
            New-Item -Path $destinationDir -ItemType Directory | Out-Null
        }

        $httpClient = [System.Net.Http.HttpClient]::new()
        $request = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Get, $DownloadUrl)
        $response = $httpClient.SendAsync($request, [System.Net.Http.HttpCompletionOption]::ResponseHeadersRead).Result
        $response.EnsureSuccessStatusCode()

        # Response Handling
        $content = $response.Content
        $totalSize = $content.Headers.ContentLength
        $fileStream = [System.IO.FileStream]::new($destinationPath, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
        $bufferSize = 8192
        $buffer = New-Object byte[] $bufferSize
        $totalRead = 0
        $progress = 0

        # Write data to file
        $stream = $content.ReadAsStreamAsync().Result
        while (($read = $stream.ReadAsync($buffer, 0, $bufferSize).Result) -gt 0) {
            $fileStream.WriteAsync($buffer, 0, $read).Wait()
            $totalRead += $read
            $percentComplete = [int](($totalRead / $totalSize) * 100)
            if ($percentComplete -gt $progress) {
                $progress = $percentComplete
                Show-Progress -PercentComplete $progress
            }
        }

        # Cleanup
        $fileStream.Close()
        $httpClient.Dispose()

        Write-Host "[i] Download completed!"
    } catch {
        Write-Error "[!] Download failed: $($_.Exception.Message)"
        exit 1
    }
}

Write-Host "OneDrive Installer [Version 1.0]"
Write-Host "Copyright (c) 2024 Wolkenhof GmbH."
Write-Host ""

# Check Server OS
$OSName = (get-computerinfo).osname
If ($OSName -like 'Microsoft Windows Server 2025*')
{
	$OS = "2025"
}
Elseif ($OSName -like 'Microsoft Windows Server 2022*')
{
	$OS = "2022"
}
Elseif ($OSName -like 'Microsoft Windows Server 2019*')
{
	$OS = "2019"
}
Elseif ($OSName -like 'Microsoft Windows Server 2016*')
{
	$OS = "2016"
}
Else
{
	$OS = "Other"
}
Write-Output "Detected Server OS: $OS"

# Create local deployment dir
If (-not (Test-Path -Path "C:\Wolkenhof\Source"))
{
	New-Item -Path "C:\Wolkenhof\Source" -ItemType "directory" -Force
}

# Download
Write-Output "[i] Downloading OneDrive ..."
DownloadOneDrive -DownloadUrl "https://go.microsoft.com/fwlink/?linkid=248256" -FileName "OneDriveSetup_x64.exe"
Unblock-File -Path "C:\Wolkenhof\source\OneDrive\OneDriveSetup_x64.exe"

Write-Output "[i] Installing OneDrive ..."
Start-Process "C:\Wolkenhof\source\OneDrive\OneDriveSetup_x64.exe" -ArgumentList "/ALLUSERS"
Do {Start-Sleep -Seconds 1} While (Get-Process -Name "OneDriveSetup_x64" -ErrorAction SilentlyContinue)
Write-Output "[i] Installation completed."