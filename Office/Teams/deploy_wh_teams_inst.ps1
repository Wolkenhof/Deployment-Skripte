<#
############################## Wolkenhof ##############################
Purpose : Install or Update Teams on Server 2019, 2022 & 2025
Created : 22.11.2024
Source  : https://github.com/Wolkenhof/Deployment-Skripte
Author  : jgu
Company : Wolkenhof GmbH
############################## Wolkenhof ##############################
#>
Add-Type -AssemblyName System.Net.Http

function DownloadTeams($DownloadUrl, $FileName, $IsClassic = $false) {
    function Show-Progress {
        param (
            [int]$PercentComplete
        )
        Write-Progress -Activity "Downloading Teams" -Status "$PercentComplete% Complete" -PercentComplete $PercentComplete
    }

    Write-Output "[i] Downloading Teams ..."
    try {
        $destinationPath = "C:\Wolkenhof\Source\Teams-New\$FileName"
        if ($IsClassic) {
            $destinationPath = "C:\Wolkenhof\Source\Teams-Classic\$FileName"
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

Write-Host "Teams Installer [Version 1.0]"
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

# Remove Teams Machine-Wide Installer
$MachineWide = Get-WmiObject -Class Win32_Product | Where-Object{$_.Name -eq "Teams Machine-Wide Installer"}
If ($MachineWide)
{
	Write-Output "[i] Removing Teams Classic (Machine-Wide) ..."
	$MachineWide.Uninstall()
}

# Remove old Teams New Meeting Add-in
$MeetingAddIn = Get-CimInstance -Class Win32_Product | Where-Object{$_.Name -eq "Microsoft Teams Meeting Add-in for Microsoft Office"}
If ($MeetingAddIn)
{
	$TMAPath = "$($MeetingAddIn.InstallSource)MicrosoftTeamsMeetingAddinInstaller.msi"
	If (Test-Path -Path $TMAPath)
	{
		$Arguments = '/x "' + $TMAPath + '" /quiet installerversion=v3 /l C:\Wolkenhof\Logs\UninstallOldTeamsAddIn.log'
		Write-Host "[i] Removing old Teams Meeting Addin ..."
		Start-Process msiexec.exe -ArgumentList $Arguments -Wait
		Start-Sleep -Seconds 15
	}
	Else
	{
		$RegPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{A7AB73A3-CB10-4AA5-9D38-6AEFFBDE4C91}"
		If (Test-Path $RegPath)
		{
			Remove-Item -Path $RegPath -Recurse -Force
            Write-Host "[i] Deleted orphaned Uninstall-Regkey!"
		}
	}
}

# Install Teams
Switch ($OS)
{
    "2025"
	{
		# Install Teams New
		Write-Output "[i] Installing Teams (New)"
		
        DownloadTeams -DownloadUrl "https://go.microsoft.com/fwlink/?linkid=2243204" -FileName "teamsbootstrapper.exe" -IsClassic $false # Bootstrapper
		Unblock-File -Path C:\Wolkenhof\source\Teams-New\teamsbootstrapper.exe

        DownloadTeams -DownloadUrl "https://go.microsoft.com/fwlink/?linkid=2196106" -FileName "MSTeams-x64.msix" -IsClassic $false # MSIX x64
		Unblock-File -Path C:\Wolkenhof\source\Teams-New\MSTeams-x64.msix
		
        If (-not (Test-Path -Path "HKLM:\SOFTWARE\Microsoft\Teams"))
		{
			New-Item -Path "HKLM:\SOFTWARE\Microsoft\Teams" -Force
		}
		Set-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\Teams -Name disableAutoUpdate -Value 1
		Start-Process C:\Wolkenhof\source\Teams-New\teamsbootstrapper.exe -wait -ArgumentList '-p -o C:\Wolkenhof\source\Teams-New\MSTeams-x64.msix'
 
        # Uninstall MeetingAddIn
        $NewTeamsPackageVersion = (Get-AppxPackage -Name MSTeams).Version
		$TMAPath = "{0}\WINDOWSAPPS\MSTEAMS_{1}_X64__8WEKYB3D8BBWE\MicrosoftTeamsMeetingAddinInstaller.MSI" -f $env:programfiles,$NewTeamsPackageVersion
        If (Test-Path -Path $TMAPath)
        {
            $Arguments = '/x "' + $TMAPath + '" /quiet installerversion=v3 /l C:\Wolkenhof\Logs\UninstallTeamsAddIn.log'
            Write-Host "[i] Uninstalling Teams Meeting Addin ..."
            Start-Process msiexec.exe -ArgumentList $Arguments -Wait
            Start-Sleep -Seconds 15
        }

		# Install MeetingAddIn
		if ($TMAVersion = (Get-AppLockerFileInformation -Path $TMAPath | Select-Object -ExpandProperty Publisher).BinaryVersion)
		{
			$TargetDir = "{0}\Microsoft\TeamsMeetingAddin\{1}\" -f ${env:ProgramFiles(x86)},$TMAVersion
			$params = '/i "{0}" TARGETDIR="{1}" /qn ALLUSERS=1' -f $TMAPath, $TargetDir
			Write-Host "[i] Installing Teams Meeting Addin ..."
			Start-Process msiexec.exe -ArgumentList $params -Wait
		}    
		break
	}

	"2022"
	{
		# Install Teams New
		Write-Output "[i] Installing Teams (New)"
		
        DownloadTeams -DownloadUrl "https://go.microsoft.com/fwlink/?linkid=2243204" -FileName "teamsbootstrapper.exe" -IsClassic $false # Bootstrapper
		Unblock-File -Path C:\Wolkenhof\source\Teams-New\teamsbootstrapper.exe

        DownloadTeams -DownloadUrl "https://go.microsoft.com/fwlink/?linkid=2196106" -FileName "MSTeams-x64.msix" -IsClassic $false # MSIX x64
		Unblock-File -Path C:\Wolkenhof\source\Teams-New\MSTeams-x64.msix
		
        If (-not (Test-Path -Path "HKLM:\SOFTWARE\Microsoft\Teams"))
		{
			New-Item -Path "HKLM:\SOFTWARE\Microsoft\Teams" -Force
		}
		Set-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\Teams -Name disableAutoUpdate -Value 1
		Start-Process C:\Wolkenhof\source\Teams-New\teamsbootstrapper.exe -wait -ArgumentList '-p -o C:\Wolkenhof\source\Teams-New\MSTeams-x64.msix'
 
        # Uninstall MeetingAddIn
        $NewTeamsPackageVersion = (Get-AppxPackage -Name MSTeams).Version
		$TMAPath = "{0}\WINDOWSAPPS\MSTEAMS_{1}_X64__8WEKYB3D8BBWE\MicrosoftTeamsMeetingAddinInstaller.MSI" -f $env:programfiles,$NewTeamsPackageVersion
        If (Test-Path -Path $TMAPath)
        {
            $Arguments = '/x "' + $TMAPath + '" /quiet installerversion=v3 /l C:\Wolkenhof\Logs\UninstallTeamsAddIn.log'
            Write-Host "[i] Uninstalling Teams Meeting Addin ..."
            Start-Process msiexec.exe -ArgumentList $Arguments -Wait
            Start-Sleep -Seconds 15
        }

		# Install MeetingAddIn
		if ($TMAVersion = (Get-AppLockerFileInformation -Path $TMAPath | Select-Object -ExpandProperty Publisher).BinaryVersion)
		{
			$TargetDir = "{0}\Microsoft\TeamsMeetingAddin\{1}\" -f ${env:ProgramFiles(x86)},$TMAVersion
			$params = '/i "{0}" TARGETDIR="{1}" /qn ALLUSERS=1' -f $TMAPath, $TargetDir
			Write-Host "[i] Installing Teams Meeting Addin ..."
			Start-Process msiexec.exe -ArgumentList $params -Wait
		}    
		break
	}
	
	"2019"
	{
		# Install Teams New
		Write-Output "[i] Installing Teams (New)"

        DownloadTeams -DownloadUrl "https://go.microsoft.com/fwlink/?linkid=2196106" -FileName "MSTeams-x64.msix" -IsClassic $false # MSIX x64
		Unblock-File -Path C:\Wolkenhof\source\Teams-New\MSTeams-x64.msix
		If (-not (Test-Path -Path "HKLM:\SOFTWARE\Microsoft\Teams"))
		{
			New-Item -Path "HKLM:\SOFTWARE\Microsoft\Teams" -Force
		}
		Set-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\Teams -Name disableAutoUpdate -Value 1
		Set-ItemProperty -Path HKLM:\SOFTWARE\Policies\Microsoft\Windows\Appx -Name AllowAllTrustedApps -Value 1
		Dism /Online /Add-ProvisionedAppxPackage /PackagePath:"C:\Wolkenhof\source\Teams-New\MSTeams-x64.msix" /SkipLicense

        # Uninstall MeetingAddIn
        $NewTeamsPackageVersion = (Get-AppxPackage -Name MSTeams).Version
		$TMAPath = "{0}\WINDOWSAPPS\MSTEAMS_{1}_X64__8WEKYB3D8BBWE\MicrosoftTeamsMeetingAddinInstaller.MSI" -f $env:programfiles,$NewTeamsPackageVersion
        If (Test-Path -Path $TMAPath)
        {
            $Arguments = '/x "' + $TMAPath + '" /quiet installerversion=v3 /l C:\Wolkenhof\Logs\UninstallTeamsAddIn.log'
            Write-Host "[i] Uninstalling Teams Meeting Addin ..."
            Start-Process msiexec.exe -ArgumentList $Arguments -Wait
            Start-Sleep -Seconds 15
        }

		# Install MeetingAddIn
		$NewTeamsPackageVersion = (Get-AppxPackage -Name MSTeams).Version
		$TMAPath = "{0}\WINDOWSAPPS\MSTEAMS_{1}_X64__8WEKYB3D8BBWE\MICROSOFTTEAMSMEETINGADDININSTALLER.MSI" -f $env:programfiles,$NewTeamsPackageVersion
		If (Test-Path -Path $TMAPath)
        {
            if ($TMAVersion = (Get-AppLockerFileInformation -Path $TMAPath | Select-Object -ExpandProperty Publisher).BinaryVersion)
		    {
			    $TargetDir = "{0}\Microsoft\TeamsMeetingAddin\{1}\" -f ${env:ProgramFiles(x86)},$TMAVersion
			    $params = '/i "{0}" TARGETDIR="{1}" /qn ALLUSERS=1' -f $TMAPath, $TargetDir
                Write-Host "[i] Installing Teams Meeting Addin ..."
			    Start-Process msiexec.exe -ArgumentList $params
            }
		}    
		break
	}

	Default
	{
		# Install Teams Classic
        Write-Output "[i] Installing Teams (Classic)"
        DownloadTeams -DownloadUrl "https://statics.teams.cdn.office.net/production-windows-x64/1.7.00.7956/Teams_windows_x64.msi" -FileName "Teams_windows_x64.msi" -IsClassic $true # MSIX x64
		Unblock-File -Path C:\Wolkenhof\source\Teams-Classic\Teams_windows_x64.msi
		Start-Process msiexec.exe -wait -ArgumentList '/i C:\Wolkenhof\source\Teams-Classic\Teams_windows_x64.msi ALLUSER=1 ALLUSERS=1'
		Remove-ItemProperty -path "hklm:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Run" -name "Teams"
		Remove-Item 'C:\Users\Public\Desktop\Microsoft Teams.lnk'
	}
}
Write-Output "[i] Teams installation completed."
# SIG # Begin signature block
# MIIocAYJKoZIhvcNAQcCoIIoYTCCKF0CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC7C8MiHwW+rRMw
# rUldcgipw+lcYc6w7QdKSAEghrXy0aCCIKQwggXJMIIEsaADAgECAhAbtY8lKt8j
# AEkoya49fu0nMA0GCSqGSIb3DQEBDAUAMH4xCzAJBgNVBAYTAlBMMSIwIAYDVQQK
# ExlVbml6ZXRvIFRlY2hub2xvZ2llcyBTLkEuMScwJQYDVQQLEx5DZXJ0dW0gQ2Vy
# dGlmaWNhdGlvbiBBdXRob3JpdHkxIjAgBgNVBAMTGUNlcnR1bSBUcnVzdGVkIE5l
# dHdvcmsgQ0EwHhcNMjEwNTMxMDY0MzA2WhcNMjkwOTE3MDY0MzA2WjCBgDELMAkG
# A1UEBhMCUEwxIjAgBgNVBAoTGVVuaXpldG8gVGVjaG5vbG9naWVzIFMuQS4xJzAl
# BgNVBAsTHkNlcnR1bSBDZXJ0aWZpY2F0aW9uIEF1dGhvcml0eTEkMCIGA1UEAxMb
# Q2VydHVtIFRydXN0ZWQgTmV0d29yayBDQSAyMIICIjANBgkqhkiG9w0BAQEFAAOC
# Ag8AMIICCgKCAgEAvfl4+ObVgAxknYYblmRnPyI6HnUBfe/7XGeMycxca6mR5rlC
# 5SBLm9qbe7mZXdmbgEvXhEArJ9PoujC7Pgkap0mV7ytAJMKXx6fumyXvqAoAl4Va
# qp3cKcniNQfrcE1K1sGzVrihQTib0fsxf4/gX+GxPw+OFklg1waNGPmqJhCrKtPQ
# 0WeNG0a+RzDVLnLRxWPa52N5RH5LYySJhi40PylMUosqp8DikSiJucBb+R3Z5yet
# /5oCl8HGUJKbAiy9qbk0WQq/hEr/3/6zn+vZnuCYI+yma3cWKtvMrTscpIfcRnNe
# GWJoRVfkkIJCu0LW8GHgwaM9ZqNd9BjuiMmNF0UpmTJ1AjHuKSbIawLmtWJFfzcV
# WiNoidQ+3k4nsPBADLxNF8tNorMe0AZa3faTz1d1mfX6hhpneLO/lv403L3nUlbl
# s+V1e9dBkQXcXWnjlQ1DufyDljmVe2yAWk8TcsbXfSl6RLpSpCrVQUYJIP4ioLZb
# MI28iQzV13D4h1L92u+sUS4Hs07+0AnacO+Y+lbmbdu1V0vc5SwlFcieLnhO+Nqc
# noYsylfzGuXIkosagpZ6w7xQEmnYDlpGizrrJvojybawgb5CAKT41v4wLsfSRvbl
# jnX98sy50IdbzAYQYLuDNbdeZ95H7JlI8aShFf6tjGKOOVVPORa5sWOd/7cCAwEA
# AaOCAT4wggE6MA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFLahVDkCw6A/joq8
# +tT4HKbROg79MB8GA1UdIwQYMBaAFAh2zcsH/yT2xc3tu5C84oQ3RnX3MA4GA1Ud
# DwEB/wQEAwIBBjAvBgNVHR8EKDAmMCSgIqAghh5odHRwOi8vY3JsLmNlcnR1bS5w
# bC9jdG5jYS5jcmwwawYIKwYBBQUHAQEEXzBdMCgGCCsGAQUFBzABhhxodHRwOi8v
# c3ViY2Eub2NzcC1jZXJ0dW0uY29tMDEGCCsGAQUFBzAChiVodHRwOi8vcmVwb3Np
# dG9yeS5jZXJ0dW0ucGwvY3RuY2EuY2VyMDkGA1UdIAQyMDAwLgYEVR0gADAmMCQG
# CCsGAQUFBwIBFhhodHRwOi8vd3d3LmNlcnR1bS5wbC9DUFMwDQYJKoZIhvcNAQEM
# BQADggEBAFHCoVgWIhCL/IYx1MIy01z4S6Ivaj5N+KsIHu3V6PrnCA3st8YeDrJ1
# BXqxC/rXdGoABh+kzqrya33YEcARCNQOTWHFOqj6seHjmOriY/1B9ZN9DbxdkjuR
# mmW60F9MvkyNaAMQFtXx0ASKhTP5N+dbLiZpQjy6zbzUeulNndrnQ/tjUoCFBMQl
# lVXwfqefAcVbKPjgzoZwpic7Ofs4LphTZSJ1Ldf23SIikZbr3WjtP6MZl9M7JYjs
# NhI9qX7OAo0FmpKnJ25FspxihjcNpDOO16hO0EoXQ0zF8ads0h5YbBRRfopUofbv
# n3l6XYGaFpAP4bvxSgD5+d2+7arszgowggaVMIIEfaADAgECAhAJxcz4u2Z9cTeq
# wVmABssxMA0GCSqGSIb3DQEBDAUAMFYxCzAJBgNVBAYTAlBMMSEwHwYDVQQKExhB
# c3NlY28gRGF0YSBTeXN0ZW1zIFMuQS4xJDAiBgNVBAMTG0NlcnR1bSBUaW1lc3Rh
# bXBpbmcgMjAyMSBDQTAeFw0yMzExMDIwODMyMjNaFw0zNDEwMzAwODMyMjNaMFAx
# CzAJBgNVBAYTAlBMMSEwHwYDVQQKDBhBc3NlY28gRGF0YSBTeXN0ZW1zIFMuQS4x
# HjAcBgNVBAMMFUNlcnR1bSBUaW1lc3RhbXAgMjAyMzCCAiIwDQYJKoZIhvcNAQEB
# BQADggIPADCCAgoCggIBALkWuurG532SNqqCQCjzkjK3p5w3fjc5Y/O004WQ5G+x
# zq6SG5w45BD6zPEfSOyLcBGMHAbVv2hDCcPHUI46Q6nCbYfNjbPG0l7ZfaoL4fwM
# y3j6dQ0BgW4wQyNF6rmm0NMjcmJ0MRuBzEp2vZrN8LCYncWmoakqvUtu0IPZjuIu
# vBk7E4OR1VgoTIkvRQ8nYDXwmA1Hnj4JnT+lV8J9s4RlqDrmjJTcDfdljzyHmaHO
# f1Yg8X+otHmq30cp727xj64yDPwwpBqAf9qNYb+5hyp5ArbwBLcSHkBxLCXjEV/A
# cZoXATHEFZJctlEZRuf1oV2KtJkop17bSnUI6WZmTEiYlj5vFBhKDDmcQzSM+Dqt
# 48P7QhBBzgA8rp1IcA5BLdC8Emt/NNaUJCiQa06/Fw0izlw69oA2ZNwZwuCQfR4e
# AwGksWVzLMTRCRjwd6H7GW1kUSIC8rmBufwIezyij2jT8mMup1ZgutbgecRLjf80
# LX+w5oJWa2yVNoWhb9ZFFu0lpGsr/TeMWOs33bV0Ke1FGKcH8TDcxDWTE83rThYI
# x4u8A6lPcXkpsFeg8Osyhb04ZNidiq/zwDqFNtUVGz4SLxQmOTgiV86ScdZ26KZE
# pDgtgNjUYNIDfdhRn9zc+ii1qdzaJY81q+PL+J4Ngh0fxdVtF9apyGcOlMT7Q0Vz
# AgMBAAGjggFjMIIBXzAMBgNVHRMBAf8EAjAAMB0GA1UdDgQWBBTHaTwu5r3jWUf/
# GRLB2TToQc/jjzAfBgNVHSMEGDAWgBS+VAIvv0Bsc0POrAklTp5DRBru4DAOBgNV
# HQ8BAf8EBAMCB4AwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwMwYDVR0fBCwwKjAo
# oCagJIYiaHR0cDovL2NybC5jZXJ0dW0ucGwvY3RzY2EyMDIxLmNybDBvBggrBgEF
# BQcBAQRjMGEwKAYIKwYBBQUHMAGGHGh0dHA6Ly9zdWJjYS5vY3NwLWNlcnR1bS5j
# b20wNQYIKwYBBQUHMAKGKWh0dHA6Ly9yZXBvc2l0b3J5LmNlcnR1bS5wbC9jdHNj
# YTIwMjEuY2VyMEEGA1UdIAQ6MDgwNgYLKoRoAYb2dwIFAQswJzAlBggrBgEFBQcC
# ARYZaHR0cHM6Ly93d3cuY2VydHVtLnBsL0NQUzANBgkqhkiG9w0BAQwFAAOCAgEA
# eN3usTpD5vfZazi9Ml4LQdwYOLuZ9BSdton2cUU5CmLM9f5gike4rz1M+Q50MXuU
# 4SZNnNVCnDSTCkhkO4HyVIzQbD0qWg69ciwaMX8qBM3FgzlpWJA0y2giIXpb3Kya
# 5sMcXuUTFJOg93Wv43TNgZeUTW4Rfij3zwr9nuTCAT8YLrj1LU4RnkgZIaaKu1yu
# 4tf/GGMgMDlL9xV/PRZ78SUdqYez5R9bf8jFOKC++rgkJt1keD0OyORb5SAYYBW2
# TEHuqKeZYlqa93CmC6MDA5PXKb+CI9NbkLz8yeQvXxmBVDfyyoqoV2pRL5khV5cp
# 9Xnwdpa1XYuKnVjSW4vsyzBvznqPPvNcg2Tv0fhd9tY6vJ/sC1YGOu6zbyOYdYre
# Bc2GPZK1Vw4jjwNzoIV9cMyj9z8T9pvbXuRNiGKG3asJZ4ZLlMdDdtlXH6VQ8toN
# 7eRVeNi/ExhApa7ThBfr69REVJ4vdZWtRI7qcSdm7tfYRhyLkxSaZR0QSIBVk7/T
# fIuU1ZQ0Zfvb/3j29T7lk32v0QZ2ntfdbuYsvVPHiAuYeesH3s7571FgrrfvQwLn
# ayK5+7XWnefw4bmzbMnDYnoukP4ctvIKB9Eh31DlQqCyPQDVC6gG63wUjph1ofex
# HWmicS/oaw1itPIG1JHvtyxRYtQLJVuiwXf5p7T5Kh8wgga5MIIEoaADAgECAhEA
# maOACiZVO2Wr3G6EprPqOTANBgkqhkiG9w0BAQwFADCBgDELMAkGA1UEBhMCUEwx
# IjAgBgNVBAoTGVVuaXpldG8gVGVjaG5vbG9naWVzIFMuQS4xJzAlBgNVBAsTHkNl
# cnR1bSBDZXJ0aWZpY2F0aW9uIEF1dGhvcml0eTEkMCIGA1UEAxMbQ2VydHVtIFRy
# dXN0ZWQgTmV0d29yayBDQSAyMB4XDTIxMDUxOTA1MzIxOFoXDTM2MDUxODA1MzIx
# OFowVjELMAkGA1UEBhMCUEwxITAfBgNVBAoTGEFzc2VjbyBEYXRhIFN5c3RlbXMg
# Uy5BLjEkMCIGA1UEAxMbQ2VydHVtIENvZGUgU2lnbmluZyAyMDIxIENBMIICIjAN
# BgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAnSPPBDAjO8FGLOczcz5jXXp1ur5c
# Tbq96y34vuTmflN4mSAfgLKTvggv24/rWiVGzGxT9YEASVMw1Aj8ewTS4IndU8s7
# VS5+djSoMcbvIKck6+hI1shsylP4JyLvmxwLHtSworV9wmjhNd627h27a8RdrT1P
# H9ud0IF+njvMk2xqbNTIPsnWtw3E7DmDoUmDQiYi/ucJ42fcHqBkbbxYDB7SYOou
# u9Tj1yHIohzuC8KNqfcYf7Z4/iZgkBJ+UFNDcc6zokZ2uJIxWgPWXMEmhu1gMXgv
# 8aGUsRdaCtVD2bSlbfsq7BiqljjaCun+RJgTgFRCtsuAEw0pG9+FA+yQN9n/kZtM
# LK+Wo837Q4QOZgYqVWQ4x6cM7/G0yswg1ElLlJj6NYKLw9EcBXE7TF3HybZtYvj9
# lDV2nT8mFSkcSkAExzd4prHwYjUXTeZIlVXqj+eaYqoMTpMrfh5MCAOIG5knN4Q/
# JHuurfTI5XDYO962WZayx7ACFf5ydJpoEowSP07YaBiQ8nXpDkNrUA9g7qf/rCkK
# bWpQ5boufUnq1UiYPIAHlezf4muJqxqIns/kqld6JVX8cixbd6PzkDpwZo4SlADa
# Ci2JSplKShBSND36E/ENVv8urPS0yOnpG4tIoBGxVCARPCg1BnyMJ4rBJAcOSnAW
# d18Jx5n858JSqPECAwEAAaOCAVUwggFRMA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0O
# BBYEFN10XUwA23ufoHTKsW73PMAywHDNMB8GA1UdIwQYMBaAFLahVDkCw6A/joq8
# +tT4HKbROg79MA4GA1UdDwEB/wQEAwIBBjATBgNVHSUEDDAKBggrBgEFBQcDAzAw
# BgNVHR8EKTAnMCWgI6Ahhh9odHRwOi8vY3JsLmNlcnR1bS5wbC9jdG5jYTIuY3Js
# MGwGCCsGAQUFBwEBBGAwXjAoBggrBgEFBQcwAYYcaHR0cDovL3N1YmNhLm9jc3At
# Y2VydHVtLmNvbTAyBggrBgEFBQcwAoYmaHR0cDovL3JlcG9zaXRvcnkuY2VydHVt
# LnBsL2N0bmNhMi5jZXIwOQYDVR0gBDIwMDAuBgRVHSAAMCYwJAYIKwYBBQUHAgEW
# GGh0dHA6Ly93d3cuY2VydHVtLnBsL0NQUzANBgkqhkiG9w0BAQwFAAOCAgEAdYhY
# D+WPUCiaU58Q7EP89DttyZqGYn2XRDhJkL6P+/T0IPZyxfxiXumYlARMgwRzLRUS
# tJl490L94C9LGF3vjzzH8Jq3iR74BRlkO18J3zIdmCKQa5LyZ48IfICJTZVJeChD
# UyuQy6rGDxLUUAsO0eqeLNhLVsgw6/zOfImNlARKn1FP7o0fTbj8ipNGxHBIutiR
# sWrhWM2f8pXdd3x2mbJCKKtl2s42g9KUJHEIiLni9ByoqIUul4GblLQigO0ugh7b
# WRLDm0CdY9rNLqyA3ahe8WlxVWkxyrQLjH8ItI17RdySaYayX3PhRSC4Am1/7mAT
# wZWwSD+B7eMcZNhpn8zJ+6MTyE6YoEBSRVrs0zFFIHUR08Wk0ikSf+lIe5Iv6RY3
# /bFAEloMU+vUBfSouCReZwSLo8WdrDlPXtR0gicDnytO7eZ5827NS2x7gCBibESY
# kOh1/w1tVxTpV2Na3PR7nxYVlPu1JPoRZCbH86gc96UTvuWiOruWmyOEMLOGGniR
# +x+zPF/2DaGgK2W1eEJfo2qyrBNPvF7wuAyQfiFXLwvWHamoYtPZo0LHuH8X3n9C
# +xN4YaNjt2ywzOr+tKyEVAotnyU9vyEVOaIYMk3IeBrmFnn0gbKeTTyYeEEUz/Qw
# t4HOUBCrW602NCmvO1nm+/80nLy5r0AZvCQxaQ4wgga5MIIEoaADAgECAhEA5/9p
# xzs1zkuRJth0fGilhzANBgkqhkiG9w0BAQwFADCBgDELMAkGA1UEBhMCUEwxIjAg
# BgNVBAoTGVVuaXpldG8gVGVjaG5vbG9naWVzIFMuQS4xJzAlBgNVBAsTHkNlcnR1
# bSBDZXJ0aWZpY2F0aW9uIEF1dGhvcml0eTEkMCIGA1UEAxMbQ2VydHVtIFRydXN0
# ZWQgTmV0d29yayBDQSAyMB4XDTIxMDUxOTA1MzIwN1oXDTM2MDUxODA1MzIwN1ow
# VjELMAkGA1UEBhMCUEwxITAfBgNVBAoTGEFzc2VjbyBEYXRhIFN5c3RlbXMgUy5B
# LjEkMCIGA1UEAxMbQ2VydHVtIFRpbWVzdGFtcGluZyAyMDIxIENBMIICIjANBgkq
# hkiG9w0BAQEFAAOCAg8AMIICCgKCAgEA6RIfBDXtuV16xaaVQb6KZX9Od9FtJXXT
# Zo7b+GEof3+3g0ChWiKnO7R4+6MfrvLyLCWZa6GpFHjEt4t0/GiUQvnkLOBRdBqr
# 5DOvlmTvJJs2X8ZmWgWJjC7PBZLYBWAs8sJl3kNXxBMX5XntjqWx1ZOuuXl0R4x+
# zGGSMzZ45dpvB8vLpQfZkfMC/1tL9KYyjU+htLH68dZJPtzhqLBVG+8ljZ1ZFilO
# KksS79epCeqFSeAUm2eMTGpOiS3gfLM6yvb8Bg6bxg5yglDGC9zbr4sB9ceIGRtC
# QF1N8dqTgM/dSViiUgJkcv5dLNJeWxGCqJYPgzKlYZTgDXfGIeZpEFmjBLwURP5A
# BsyKoFocMzdjrCiFbTvJn+bD1kq78qZUgAQGGtd6zGJ88H4NPJ5Y2R4IargiWAmv
# 8RyvWnHr/VA+2PrrK9eXe5q7M88YRdSTq9TKbqdnITUgZcjjm4ZUjteq8K331a4P
# 0s2in0p3UubMEYa/G5w6jSWPUzchGLwWKYBfeSu6dIOC4LkeAPvmdZxSB1lWOb9H
# zVWZoM8Q/blaP4LWt6JxjkI9yQsYGMdCqwl7uMnPUIlcExS1mzXRxUowQref/EPa
# S7kYVaHHQrp4XB7nTEtQhkP0Z9Puz/n8zIFnUSnxDof4Yy650PAXSYmK2TcbyDoT
# Nmmt8xAxzcMCAwEAAaOCAVUwggFRMA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYE
# FL5UAi+/QGxzQ86sCSVOnkNEGu7gMB8GA1UdIwQYMBaAFLahVDkCw6A/joq8+tT4
# HKbROg79MA4GA1UdDwEB/wQEAwIBBjATBgNVHSUEDDAKBggrBgEFBQcDCDAwBgNV
# HR8EKTAnMCWgI6Ahhh9odHRwOi8vY3JsLmNlcnR1bS5wbC9jdG5jYTIuY3JsMGwG
# CCsGAQUFBwEBBGAwXjAoBggrBgEFBQcwAYYcaHR0cDovL3N1YmNhLm9jc3AtY2Vy
# dHVtLmNvbTAyBggrBgEFBQcwAoYmaHR0cDovL3JlcG9zaXRvcnkuY2VydHVtLnBs
# L2N0bmNhMi5jZXIwOQYDVR0gBDIwMDAuBgRVHSAAMCYwJAYIKwYBBQUHAgEWGGh0
# dHA6Ly93d3cuY2VydHVtLnBsL0NQUzANBgkqhkiG9w0BAQwFAAOCAgEAuJNZd8lM
# Ff2UBwigp3qgLPBBk58BFCS3Q6aJDf3TISoytK0eal/JyCB88aUEd0wMNiEcNVMb
# K9j5Yht2whaknUE1G32k6uld7wcxHmw67vUBY6pSp8QhdodY4SzRRaZWzyYlviUp
# yU4dXyhKhHSncYJfa1U75cXxCe3sTp9uTBm3f8Bj8LkpjMUSVTtMJ6oEu5JqCYzR
# fc6nnoRUgwz/GVZFoOBGdrSEtDN7mZgcka/tS5MI47fALVvN5lZ2U8k7Dm/hTX8C
# WOw0uBZloZEW4HB0Xra3qE4qzzq/6M8gyoU/DE0k3+i7bYOrOk/7tPJg1sOhytOG
# UQ30PbG++0FfJioDuOFhj99b151SqFlSaRQYz74y/P2XJP+cF19oqozmi0rRTkfy
# EJIvhIZ+M5XIFZttmVQgTxfpfJwMFFEoQrSrklOxpmSygppsUDJEoliC05vBLVQ+
# gMZyYaKvBJ4YxBMlKH5ZHkRdloRYlUDplk8GUa+OCMVhpDSQurU6K1ua5dmZftnv
# SSz2H96UrQDzA6DyiI1V3ejVtvn2azVAXg6NnjmuRZ+wa7Pxy0H3+V4K4rOTHlG3
# VYA6xfLsTunCz72T6Ot4+tkrDYOeaU1pPX1CBfYj6EW2+ELq46GP8KCNUQDirWLU
# 4nOmgCat7vN0SD6RlwUiSsMeCiQDmZwgwrUwggbAMIIEqKADAgECAhAsNLzevp4x
# j/t17DH2xcVtMA0GCSqGSIb3DQEBCwUAMFYxCzAJBgNVBAYTAlBMMSEwHwYDVQQK
# ExhBc3NlY28gRGF0YSBTeXN0ZW1zIFMuQS4xJDAiBgNVBAMTG0NlcnR1bSBDb2Rl
# IFNpZ25pbmcgMjAyMSBDQTAeFw0yNDA4MTIwNjAzNDZaFw0yNTA4MTIwNjAzNDVa
# MGYxCzAJBgNVBAYTAkRFMRYwFAYDVQQIDA1OaWVkZXJzYWNoc2VuMQ8wDQYDVQQH
# DAZVZWx6ZW4xFjAUBgNVBAoMDUpvbmFzIEfDvG5uZXIxFjAUBgNVBAMMDUpvbmFz
# IEfDvG5uZXIwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQCya04+xjpQ
# +XGrtmvuKnwqMiHCcy9sk5FQ7GRx99e3DmwHDlZWyWDt8gt+oRGah+AZ4Xy8Z6WX
# Mmb2v+RW55TOQgDSUBL+IwgXjcyHK2R1rGT21zJ5Zk+J/iDz7xRB0pgkOxe+Qm9i
# ZiIYTYzMQZ5ZQK4zW1cAk9qZwQJHoOMucQvj6gOkPq4iFMcNN7S2nhCupMQ//NFd
# LGcMIKCpwVZgU03WjcJxVuklMsC1mluZYht1EY9uZTgKhrTtXGNUmW67bvKCDoJ6
# 03OYPsBUxyozgEk2GnZYlJF2VhywYYE2t1cXlqynzVwaxpnQqcRF0/IQJQIalVU/
# Pq+NrZP5fOUBuWWpoK/ZdBtcW+YUeR0ZSGxK5ECitzNtuSMh9N+3nLqGiUplkvHi
# cl10sxqJhEICjmkBoCXfQOMdKBePIzwAdFeWoKZFMS5pIIYfL09s1BsqWiCo/mLs
# O8VAsVzqs7QJPIsUPlcd2IaN3ok39bKnrr7s/zoVwb305cwi8LT8bZ30sAOuvEOP
# EEtwpt+/2iQ5KNs/FlcjHWopucrmfx+Uk/7lY1hRzgtXAORHwfn3B0UZ/yrAp2cB
# Echk7joEqBDBG1/cj5lI71E7XQ50VKW2uULXJOdDGqDM4gRJjYkXETixX41vFgZE
# NPkmPm9KFdxs+/ucdmD58eMZZtCHW9nO+wIDAQABo4IBeDCCAXQwDAYDVR0TAQH/
# BAIwADA9BgNVHR8ENjA0MDKgMKAuhixodHRwOi8vY2NzY2EyMDIxLmNybC5jZXJ0
# dW0ucGwvY2NzY2EyMDIxLmNybDBzBggrBgEFBQcBAQRnMGUwLAYIKwYBBQUHMAGG
# IGh0dHA6Ly9jY3NjYTIwMjEub2NzcC1jZXJ0dW0uY29tMDUGCCsGAQUFBzAChilo
# dHRwOi8vcmVwb3NpdG9yeS5jZXJ0dW0ucGwvY2NzY2EyMDIxLmNlcjAfBgNVHSME
# GDAWgBTddF1MANt7n6B0yrFu9zzAMsBwzTAdBgNVHQ4EFgQU8OWGpSKi6IdJe4FG
# ovFx1FwCGckwSwYDVR0gBEQwQjAIBgZngQwBBAEwNgYLKoRoAYb2dwIFAQQwJzAl
# BggrBgEFBQcCARYZaHR0cHM6Ly93d3cuY2VydHVtLnBsL0NQUzATBgNVHSUEDDAK
# BggrBgEFBQcDAzAOBgNVHQ8BAf8EBAMCB4AwDQYJKoZIhvcNAQELBQADggIBAERR
# g1+lEi84kIi1kYsecQ+OUMBMKdVzBIyYZ//gSyWpNIWlnEcNxxXfvCLBrAThF18D
# dUGyxnIDw2ZG3e3qglAWpeTdiMOhbpdaMI/fSKNH2/PJ//vUlD6aD0tClPn87JpH
# Bqf81GwvwCBDhFeHW5VdwMyxto/ABNqweB5sKF1wMTb0qIG1lis9wpo0AVVPXSaW
# 17JPKuGF1fNOMxxrZBUyZENxb2JcneX5jMAV9W6GSdp/U2n5OhHpIXPQjVvp5wC/
# zXqNu7zK1k/tfK8WDTpvT5RIS07d7Lr8v/R4csG4x1cN1q65n5O5j+NotVD3bDUj
# OHpenzAKmdvMCWiiHzEvz/2lExrG0eZM6ZNo/cjyUlHx4CZYFfVUeG+/IEkUb90m
# mxrvLcvEmAWYN6VqzF4v0ulA1Ye6fzJR0nnPUOf40kYI2nKa8tzKL2SCpn11sVyC
# U93Pl1iK10Ku/GlhWbmv5ULRQTND0ihK7B/x64JqUxlGtMs4mvP/Hjl/BJbpLRvC
# ueoR0+iINWxhiaUFHdYZR6t9oGQcUHGvkwwG4RGPIDDeGFhfGW/DoAqpyKSRO/hK
# 7X15Ir2ibRQzdwdDCZ0poT4wBZDswRUYunSlsDiHLrbE851zkqths+Olr01Yg9Yh
# XLvy8I48InuTL7SQS+LGA/hSBRZcKFxZWcRlCjvbMYIHIjCCBx4CAQEwajBWMQsw
# CQYDVQQGEwJQTDEhMB8GA1UEChMYQXNzZWNvIERhdGEgU3lzdGVtcyBTLkEuMSQw
# IgYDVQQDExtDZXJ0dW0gQ29kZSBTaWduaW5nIDIwMjEgQ0ECECw0vN6+njGP+3Xs
# MfbFxW0wDQYJYIZIAWUDBAIBBQCggYQwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKA
# ADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYK
# KwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQgWJw78V5v/4ivJKoW5XWhBd8eZOoF
# O7kHwRWMHkcKqp0wDQYJKoZIhvcNAQEBBQAEggIAhM837Q5JOBhxNBqGKeq0QVJk
# 0opSIdy98Qbmy3ME0KbLjCBMRrmbaYBagZL3K4nDSPbepyyeRBJlUnZ4b0/OKh9a
# 6dOlZCGRqXQrIAszuIUCEULF7Ozae0XsqSLwZoc5fS82TlsDncih9/e7fLfndMXn
# jmc5kDvSA9VCPZdsZkmyEf2q547ErIogQqzLrOGs5VwIRoL06T2cBtlehPFL1yZ+
# dii5OKZpAlTHYm5D8sz4blZW45jajv0+mS7hmiU5RbbOwbMWtuqbeRAUu1GQI5Uk
# JaowyU4659qRMXZn9xO8c+P+CQkcCDxna6I65EwotZxJ6s83xutbeZA3lWqVKxfl
# YxeNg3S8t3B9nQtELXpOK7NwWhjhaFBrnEkVm0dhvUjXtAuPxERs8g253G9FtIY/
# QhSiGmgpo+uhxb8ZzV8kT83q/xHWVYzi8/QatPByOnn8imSiYdEV5272xk37i5Fl
# MYynaE0XVekKK8Kf0lNjG8086aQ0h11fHo15TrMyhkD8jNAb0cPOLBEoINCZC5Py
# TWgHPG1MOEreSVZwhVOZUgehQjdB/maq2EbaYPH8MRn+Sx0IVg3VpnKqRmam3MaG
# U1XROtPwYGFOvN3hlG+IC8ypC/L7ToClIs8NSxR+jRQ3C18MMcpGzx/1ixkJH36Y
# JipMoDfLneoxfECRR6OhggQCMIID/gYJKoZIhvcNAQkGMYID7zCCA+sCAQEwajBW
# MQswCQYDVQQGEwJQTDEhMB8GA1UEChMYQXNzZWNvIERhdGEgU3lzdGVtcyBTLkEu
# MSQwIgYDVQQDExtDZXJ0dW0gVGltZXN0YW1waW5nIDIwMjEgQ0ECEAnFzPi7Zn1x
# N6rBWYAGyzEwDQYJYIZIAWUDBAICBQCgggFWMBoGCSqGSIb3DQEJAzENBgsqhkiG
# 9w0BCRABBDAcBgkqhkiG9w0BCQUxDxcNMjQxMTIyMTU1MjQxWjA3BgsqhkiG9w0B
# CRACLzEoMCYwJDAiBCDqlUux0EC0MUBI2GWfj2FdiHQszOBnkuBWAk1LADrTHDA/
# BgkqhkiG9w0BCQQxMgQwKPwIUbVk28e8ZfKXiTFqrgM1sqHxGrVrfDtbw/VLvtuF
# QGhd+yQ2YNS7i/26e1trMIGfBgsqhkiG9w0BCRACDDGBjzCBjDCBiTCBhgQUD0+4
# VR7/2Pbef2cmtDwT0Gql53cwbjBapFgwVjELMAkGA1UEBhMCUEwxITAfBgNVBAoT
# GEFzc2VjbyBEYXRhIFN5c3RlbXMgUy5BLjEkMCIGA1UEAxMbQ2VydHVtIFRpbWVz
# dGFtcGluZyAyMDIxIENBAhAJxcz4u2Z9cTeqwVmABssxMA0GCSqGSIb3DQEBAQUA
# BIICAK65YGXU7ZFyNJxaDPADGZXFRjxxyU7yK6rsddUHBVHR8JIC8x07uBpw6WKq
# toP6STMfX+JV7LmqVOzpeKEBCr5EtMxd58n0v62qfa7rVb813+4Edrr6WrOH381V
# stb0QA6f0T5AucDyCMWqporKaMXZ4UlpXY+K5ARON79saFNXHSSNaXtMKHNr8h+9
# oI7IuW4767ZlydjBwq+ZiQevkc+ssbMTHqiUhAcUzBv6JIqBdxH2v4bhQdGD64Pk
# 4ubracif6ge7UC54BHZsHauknDNx4D87jvicyQgPtdDJmiVpfGJksYK1TeIsuRqb
# t/hPLstF6KX5wi81QzHwuZyjpV5jXfQhK/I11KZltgxqgZ3KxdF1spP9pLuOOL9Z
# YXT7iXX/f5NftXVMrnr8VVbyx9WIsXIInHvJTC34WKxAVXGdFAT0jslJnLD1nB1H
# 9qu5fJR2eesbzXfsdVsLcigB43VA/vqbA9Wf4rryYqdah8IbSkx5rkxB4Qe9tFJ1
# IfhipixiSYRQvsgRiKVCVTXmWyBMS7YkavYK3Fid/yFOHYjd2FvBv1UyqGuweYl6
# m6n1+18/kjPbrvAXeN03rVuncehR33JgkvBNYl3YJzoJqmE7RFHAN+UBWtc1x+Sw
# 8N3rXHSugCFK8mSZExzsVkZAJjD/D7qEcn3xR6BrFpXp/sbX
# SIG # End signature block
