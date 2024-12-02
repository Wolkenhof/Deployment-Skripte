<#
############################## Wolkenhof ##############################
Purpose : FSLogix Installationsscript for Domain Controller (Main)
Created : 27.11.2024
Source  : https://github.com/Wolkenhof/Deployment-Skripte
Author  : jgu
Company : Wolkenhof GmbH
############################## Wolkenhof ##############################
#>

Write-Host "FSLogix Installationsscript (001) [Version 1.0]"
Write-Host "Copyright (c) 2024 Wolkenhof GmbH."
Write-Host ""
$script002Name = "002_fslogix_ts.ps1"

### Script Begin ###
function Show-Progress {
	param (
		[int]$PercentComplete
	)
	Write-Progress -Activity "Lade FSLogix herunter" -Status "$PercentComplete% abgeschlossen" -PercentComplete $PercentComplete
}

### Abfragen ###
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Net.Http

$FSLogixStore = "D:"
$FslogixHost = "$env:COMPUTERNAME"

### WPF Logic ###
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="FSLogix Server-Installation | Wolkenhof GmbH" Height="150" Width="460" ResizeMode="NoResize">
    <Grid>
        <Label Content="Wo soll der FSLogix Store erstellt werden?" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top"/>
        <ComboBox x:Name="DriveComboBox" Width="420" Height="30" VerticalAlignment="Top" HorizontalAlignment="Left" Margin="15,41,0,0"/>
        <Button Content="OK" x:Name="OK_Button" HorizontalAlignment="Left" Margin="360,84,0,0" VerticalAlignment="Top" Width="75"/>
        <Button Content="Abbrechen" x:Name="Cancel_Button" HorizontalAlignment="Left" Margin="280,84,0,0" VerticalAlignment="Top" Width="75"/>
    </Grid>
</Window>
"@
$reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]$xaml)
$window = [System.Windows.Markup.XamlReader]::Load($reader)

# Button Events
$okButton = $window.FindName("OK_Button")
$cancelButton = $window.FindName("Cancel_Button")
$comboBox = $window.FindName("DriveComboBox")

$okButton.Add_Click({
    $window.Tag = @{
        TextBox1Value = $textBox1.Text
        TextBox2Value = $textBox2.Text
    }
    $window.DialogResult = $true
    $window.Close()
})

$cancelButton.Add_Click({
    $window.DialogResult = $false
    $window.Close()
})

####################################################################################

$xamlInstallFsLogix = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="FSLogix Server-Installation | Wolkenhof GmbH" Height="140" Width="300" ResizeMode="NoResize">
    <Grid>
        <Label Content="Soll FSLogix jetzt installiert werden?" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top"/>
        <Button Content="Nein" x:Name="Cancel_Button" HorizontalAlignment="Left" Margin="195,67,0,0" VerticalAlignment="Top" Width="75"/>
        <Button Content="Ja" x:Name="OK_Button" HorizontalAlignment="Left" Margin="109,67,0,0" VerticalAlignment="Top" Width="75"/>
    </Grid>
</Window>
"@
$readerInstallFsLogix = [System.Xml.XmlReader]::Create([System.IO.StringReader]$xamlInstallFsLogix)
$windowInstallFsLogix = [System.Windows.Markup.XamlReader]::Load($readerInstallFsLogix)

# Button Events
$okButtonInstallFsLogix = $windowInstallFsLogix.FindName("OK_Button")
$cancelButtonInstallFsLogix = $windowInstallFsLogix.FindName("Cancel_Button")

$okButtonInstallFsLogix.Add_Click({
    $windowInstallFsLogix.DialogResult = $true
    $windowInstallFsLogix.Close()
})

$cancelButtonInstallFsLogix.Add_Click({
    $windowInstallFsLogix.DialogResult = $false
    $windowInstallFsLogix.Close()
})

####################################################################################

# GPO Abfrage
$GroupOU = Get-ADOrganizationalUnit -Filter * | Select-Object Name, DistinguishedName | Out-GridView -Title "In welcher OU sollen die beiden Gruppen angelegt werden?" -PassThru
if($null -eq $GroupOU) {
    Write-Host "Abbrechen..."
	exit 1
}
$GroupOU = Get-ADOrganizationalUnit $GroupOU.DistinguishedName
$RDSHOU = Get-ADOrganizationalUnit -Filter * | Select-Object Name, DistinguishedName | Out-GridView -Title "Wo soll die FSLogix-GPO erstellt werden?" -PassThru
if($null -eq $RDSHOU) {
    Write-Host "Abbrechen..."
	exit 1
}
$RDSHOU = Get-ADOrganizationalUnit $RDSHOU.DistinguishedName

# FSLogix Store abfrage

$FSLStore = Get-ADComputer -Filter "OperatingSystem -like '*Windows Server*'" | Select-Object Name | Out-GridView -Title "Auf welchem Server soll die FSLogix-Freigabe erstellt werden?" -PassThru
if($null -eq $FSLStore) {
    Write-Host "Abbrechen..."
	exit 1
}
$FSLStoreServer = $FSLStore.Name
Write-Host "Verbindung mit Server '$FSLStoreServer' wird hergestellt ..." -ForegroundColor Yellow

$session = New-PSSession -ComputerName $FSLStoreServer
$scriptBlock = {
    $drives = [System.IO.DriveInfo]::GetDrives()
    $drives | ForEach-Object { $_.Name }
}
$driveNames = Invoke-Command -Session $session -ScriptBlock $scriptBlock
foreach ($driveName in $driveNames) {
    $comboBox.Items.Add($driveName)
}
$comboBox.SelectedIndex = 0

# Window options
$window.Owner = [System.Windows.Application]::Current.MainWindow
$result = $window.ShowDialog()

# Auswertung der Abfrage
if ($result -eq $true) {
    $FSLogixStore = $window.FindName("DriveComboBox").Text + "FSLogixStore"
    Write-Output "FSLogix Store Pfad: $FSLogixStore"
	Invoke-Command -Session $session -ScriptBlock {
		param($fsLogixStorePath)
		### Berechtigungen & Freigaben ###
		Enable-NetFirewallRule -Name FPS-SMB-In-TCP, FPS-SMB-Out-TCP
		Write-Host "Firewall Regel wurde gesetzt."  -ForegroundColor Green

		# Freigabe: FSLogixStore 
		$FSLogixShare = "FSLogixStore$" 
		$IcaclsPath = "C:\Windows\System32\icacls.exe"
		Write-Host "Erstelle $FSLogixShare -> $fsLogixStorePath ..."
		if(!(Test-Path $fsLogixStorePath)){
			New-Item -Path $fsLogixStorePath -ItemType Directory | Out-Null
		} else{
			Write-Host "ACHTUNG: Berechtigungen für ""$fsLogixStorePath"" wird zurückgesetzt!" -ForegroundColor Red
			net share /del $FSLogixShare
		}
		Start-Process -FilePath $IcaclsPath -ArgumentList $($fsLogixStorePath + " /inheritance:d /T /C /Q") -Wait
		Start-Process -FilePath $IcaclsPath -ArgumentList $($fsLogixStorePath + " /remove:g *S-1-3-0 /remove:g *S-1-1-0 /remove:g *S-1-5-18 /remove:g *S-1-5-32-545 /grant:r *S-1-5-32-544:(OI)(CI)(F) /grant:r *S-1-5-32-545:(M) /grant:r *S-1-3-0:(OI)(CI)(IO)(M) /T /C /Q") -Wait
		if(!(Get-SmbShare | ? Name -eq $FSLogixShare)){
			Write-Host "Freigabe ""$FSLogixShare"" (""$fsLogixStorePath"") existiert bislang nicht. Erstelle neue Freigabe ..."
			New-SmbShare -Name $FSLogixShare -Path $fsLogixStorePath -ChangeAccess *S-1-5-11 -CachingMode None -Description "FSLogix Profildaten" | Out-Null
		} else {
			Write-Host "ACHTUNG: Freigabe ""$FSLogixShare"" wird zurückgesetzt!" -ForegroundColor Red
			Start-Sleep 5
			net share /del $FSLogixShare
			exit 1
		}
	} -ArgumentList $FSLogixStore
} else {
    Write-Output "Der Vorgang wurde abgebrochen."
	exit 1
}

####################################################################################

# Freigabe: Install
$Install = "C:\Temp\fslogix\install"
$FSLogixShare = "FXLogix_Install$"
$IcaclsPath = "C:\Windows\System32\icacls.exe"
Write-Host "Erstelle $FSLogixShare -> $Install ..."

if(!(Test-Path $Install)){
	New-Item -Path $Install -ItemType Directory | Out-Null
} else{
	Write-Host "ACHTUNG: Berechtigungen für ""$Install"" wird zurückgesetzt!" -ForegroundColor Red
}
Start-Process -FilePath $IcaclsPath -ArgumentList $($Install + " /inheritance:d /T /C /Q") -Wait
Start-Process -FilePath $IcaclsPath -ArgumentList $($Install + " /remove:g *S-1-3-0 /remove:g *S-1-1-0 /remove:g *S-1-5-18 /remove:g *S-1-5-32-545 /grant:r *S-1-5-32-544:(OI)(CI)(F) /grant:r *S-1-5-32-545:(M) /grant:r *S-1-3-0:(OI)(CI)(IO)(M) /T /C /Q") -Wait
if(!(Get-SmbShare | ? Name -eq $FSLogixShare)){
	New-SmbShare -Name $FSLogixShare -Path $Install -ChangeAccess *S-1-5-11 -CachingMode None -Description "FSLogix Setup" | Out-Null
} else {
	Write-Host "WARNUNG: Freigabe ""$FSLogixShare"" existiert bereits! Abbrechen..."  -ForegroundColor Yellow
	exit 1
}

Write-Host "Freigaben wurden erfolgreich erstellt."  -ForegroundColor Green
####################################################################################

### Download FSLogix ###
$FslogixUrl= "https://aka.ms/fslogix_download"

# Script begin
$value = 0.0
$filePath = "C:\Temp\fslogix\install\x64\Release\FSLogixAppsSetup.exe"
if (Test-Path %filePath) {
	$value = (Get-Item $filePath).VersionInfo.FileVersion
}

# Neuste Version
try {
	Write-Host "Suche nach neue FSLogix Version ..."
    $httpClient = New-Object System.Net.Http.HttpClient
	$requestMessage = New-Object System.Net.Http.HttpRequestMessage 'HEAD', $FslogixUrl
	$response = $httpClient.SendAsync($requestMessage).GetAwaiter().GetResult()
	if ($response.IsSuccessStatusCode) {
		$finalUrl = $response.RequestMessage.RequestUri.AbsoluteUri
		$fileName = [System.IO.Path]::GetFileName($finalUrl)
		$regexPattern = "\d+\.\d+\.\d+\.\d+"
		$neueVersion = 0.0

		if ($fileName -match $regexPattern) {
			$neueVersion = $matches[0]
		}

		Write-Host "Aktuelle Version: $value.$propertyName"
		Write-Host "Neue Version: $neueVersion"
		Write-Host ""
	} else {
		Write-Error "Fehler bei der Webanfrage: $($response.ReasonPhrase)"
		exit 1
	}
}
catch {
	Write-Error "Fehler bei der Webanfrage: $_.Exception.Message"
	exit 1
}

if ($value.$propertyName -eq $neueVersion) {
	$nachricht = "FSLogix ist aktuell. Überspringe Download ..."
	Write-Host $nachricht
}
else {
	$nachricht = "Eine neue Version ist verfügbar: " + $neueVersion
	Write-Host $nachricht
	
	# Die literalen Pfade sind Absicht, anders waere RDS auch nicht supported
	Remove-Item "C:\Temp\fslogix\install" -Force -Recurse
	mkdir "C:\Temp\fslogix\install" -Force | Out-Null
	Write-Host "Download von $FslogixUrl ..."
	
	try {
		$destinationPath = "C:\Temp\fslogix\install\FSLogixAppsSetup.zip"
		
		$destinationDir = [System.IO.Path]::GetDirectoryName($destinationPath)
		if (-not (Test-Path -Path $destinationDir)) {
			New-Item -Path $destinationDir -ItemType Directory | Out-Null
		}

		$httpClient = [System.Net.Http.HttpClient]::new()
		$request = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Get, $FslogixUrl)
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

		Write-Host "Download abgeschlossen!"
	} catch {
		Write-Error "Fehler bei der Webanfrage: $($_.Exception.Message)"
		exit 1
	}
	# Auspacken hierher
	Write-Host "Auspacken von FSLogixAppsSetup.zip ..."
	Expand-Archive -LiteralPath "C:\Temp\fslogix\install\FSLogixAppsSetup.zip" -DestinationPath "C:\Temp\fslogix\install" -Force -Verbose
}
Write-Host "FSLogix wurde erfolgreich heruntergeladen."  -ForegroundColor Green

####################################################################################

### GPO Anlegung ###

$FSLogixShare = "FSLogixStore$"
$FSLogixInstallShare = "FXlogix_Install$"
$FSLogixSetup = "\\$FslogixHost\$FSLogixInstallShare\"
$FSLogixO365Group = "grp_O365_Container"
$FSLogixProfileGroup = "grp_Profile_Container"
$FSLogixPolicy = "C_FSLogix_Config"
 
Write-Host "Kopiere ADML & ADMX Dateien ..."
if(!(Test-Path \\$env:UserDNSDomain\sysvol\$env:UserDNSDomain\Policies\)){
    Copy-Item C:\Windows\PolicyDefinitions \\$env:UserDNSDomain\sysvol\$env:UserDNSDomain\Policies\ -Recurse
}
Copy-Item $("C:\Temp\fslogix\install\fslogix.adml") $env:windir\PolicyDefinitions\en-US\
Copy-Item $("C:\Temp\fslogix\install\fslogix.adml") $env:windir\PolicyDefinitions\de-DE\
Copy-Item $("C:\Temp\fslogix\install\fslogix.admx") $env:windir\PolicyDefinitions\
Write-Host "Dateien wurden erfolgreich kopiert."  -ForegroundColor Green

New-ADGroup -Name $FSLogixO365Group -DisplayName $FSLogixO365Group -Description "FSLogix Office 365 Container User" -Path $GroupOU.DistinguishedName -GroupCategory Security -GroupScope DomainLocal
New-ADGroup -Name $FSLogixProfileGroup -DisplayName $FSLogixProfileroup -Description "FSLogix Profile Container User" -Path $GroupOU.DistinguishedName -GroupCategory Security -GroupScope DomainLocal
Write-Host "AD Gruppen wurden erfolgreich angelegt."  -ForegroundColor Green

$GPOObject = New-GPO $FSLogixPolicy | New-GPLink -Target $RDSHOU.DistinguishedName -LinkEnabled Yes
 
# Profile Container
Set-GPRegistryValue -Name $FSLogixPolicy -Key "HKLM\Software\FSLogix\Profiles" -ValueName "Enabled" -Value 1 -Type DWord | Out-Null
Set-GPRegistryValue -Name $FSLogixPolicy -Key "HKLM\SOFTWARE\FSLogix\Profiles" -ValueName "RoamIdentity" -Value 1 -Type DWord | Out-Null
Set-GPRegistryValue -Name $FSLogixPolicy -Key "HKLM\Software\FSLogix\Profiles" -ValueName "IsDynamic" -Value 1 -Type DWord | Out-Null
Set-GPRegistryValue -Name $FSLogixPolicy -Key "HKLM\Software\FSLogix\Profiles" -ValueName "SizeInMBs" -Value 5000 -Type DWord | Out-Null
Set-GPRegistryValue -Name $FSLogixPolicy -Key "HKLM\Software\FSLogix\Profiles" -ValueName "RoamSearch" -Value 1 -Type DWord | Out-Null
Set-GPRegistryValue -Name $FSLogixPolicy -Key "HKLM\Software\FSLogix\Profiles" -ValueName "VHDLocations" -Value $("\\" + $FSLStoreServer + "\" + $FSLogixShare) -Type String | Out-Null
Set-GPRegistryValue -Name $FSLogixPolicy -Key "HKLM\Software\FSLogix\Profiles" -ValueName "VolumeType" -Value "VHDX" -Type String | Out-Null
 
# Office Container
Set-GPRegistryValue -Name $FSLogixPolicy -Key "HKLM\Software\Policies\FSLogix\ODFC" -ValueName "Enabled" -Value 1 -Type DWord | Out-Null
Set-GPRegistryValue -Name $FSLogixPolicy -Key "HKLM\Software\Policies\FSLogix\ODFC" -ValueName "IsDynamic" -Value 1 -Type DWord | Out-Null
Set-GPRegistryValue -Name $FSLogixPolicy -Key "HKLM\Software\Policies\FSLogix\ODFC" -ValueName "SizeInMBs" -Value 5000 -Type DWord | Out-Null
Set-GPRegistryValue -Name $FSLogixPolicy -Key "HKLM\Software\Policies\FSLogix\ODFC" -ValueName "IncludeSkype" -Value 1 -Type DWord | Out-Null
Set-GPRegistryValue -Name $FSLogixPolicy -Key "HKLM\Software\Policies\FSLogix\ODFC" -ValueName "IncludeTeams" -Value 1 -Type DWord | Out-Null
Set-GPRegistryValue -Name $FSLogixPolicy -Key "HKLM\Software\Policies\FSLogix\ODFC" -ValueName "IncludeOneNote" -Value 1 -Type DWord | Out-Null
Set-GPRegistryValue -Name $FSLogixPolicy -Key "HKLM\Software\Policies\FSLogix\ODFC" -ValueName "IncludeOutlook" -Value 1 -Type DWord | Out-Null
Set-GPRegistryValue -Name $FSLogixPolicy -Key "HKLM\Software\Policies\FSLogix\ODFC" -ValueName "IncludeOfficeActivation" -Value 1 -Type DWord | Out-Null
Set-GPRegistryValue -Name $FSLogixPolicy -Key "HKLM\Software\Policies\FSLogix\ODFC" -ValueName "IncludeOneDrive" -Value 1 -Type DWord | Out-Null
Set-GPRegistryValue -Name $FSLogixPolicy -Key "HKLM\Software\Policies\FSLogix\ODFC" -ValueName "IncludeSharepoint" -Value 1 -Type DWord | Out-Null
Set-GPRegistryValue -Name $FSLogixPolicy -Key "HKLM\Software\Policies\FSLogix\ODFC" -ValueName "IncludeOfficeFileCache" -Value 1 -Type DWord | Out-Null
Set-GPRegistryValue -Name $FSLogixPolicy -Key "HKLM\Software\Policies\FSLogix\ODFC" -ValueName "IncludeOutlookPersonalization" -Value 1 -Type DWord | Out-Null
Set-GPRegistryValue -Name $FSLogixPolicy -Key "HKLM\Software\Policies\FSLogix\ODFC" -ValueName "IncludeOneNote_UWP" -Value 1 -Type DWord | Out-Null
Set-GPRegistryValue -Name $FSLogixPolicy -Key "HKLM\Software\Policies\FSLogix\ODFC" -ValueName "OutlookCachedMode" -Value 1 -Type DWord | Out-Null
Set-GPRegistryValue -Name $FSLogixPolicy -Key "HKLM\Software\Policies\FSLogix\ODFC" -ValueName "RoamSearch" -Value 2 -Type DWord | Out-Null
Set-GPRegistryValue -Name $FSLogixPolicy -Key "HKLM\Software\Policies\FSLogix\ODFC" -ValueName "VHDLocations" -Value $("\\" + $FSLStoreServer + "\" + $FSLogixShare) -Type String | Out-Null
Set-GPRegistryValue -Name $FSLogixPolicy -Key "HKLM\Software\Policies\FSLogix\ODFC" -ValueName "VolumeType" -Value "VHDX" -Type String | Out-Null
 
# Group Policy Preferences Gruppen
New-Item $("\\" + $env:USERDNSDOMAIN + "\SYSVOL\" + $env:USERDNSDOMAIN + "\Policies\{" + $GPOObject.GpoId + "}\Machine\Preferences\Groups\Groups.xml") -ItemType File -Force | Out-Null
"<?xml version=""1.0"" encoding=""utf-8""?>" | Out-File $("\\" + $env:USERDNSDOMAIN + "\SYSVOL\" + $env:USERDNSDOMAIN + "\Policies\{" + $GPOObject.GpoId + "}\Machine\Preferences\Groups\Groups.xml") -Encoding utf8
"<Groups clsid=""{3125E937-EB16-4b4c-9934-544FC6D24D26}""><Group clsid=""{6D4A79E4-529C-4481-ABD0-F5BD7EA93BA7}"" name=""FSLogix ODFC Include List"" image=""2"" changed=""$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")"" uid=""{7D296515-BA25-452D-85A5-B1DB2580B2D5}""><Properties action=""U"" newName="""" description=""Members of this group are on the include list for Outlook Data Folder Containers"" deleteAllUsers=""1"" deleteAllGroups=""1"" removeAccounts=""0"" groupSid="""" groupName=""FSLogix ODFC Include List""><Members><Member name=""$FSLogixO365Group"" action=""ADD"" sid=""""/></Members></Properties></Group>" | Out-File $("\\" + $env:USERDNSDOMAIN + "\SYSVOL\" + $env:USERDNSDOMAIN + "\Policies\{" + $GPOObject.GpoId + "}\Machine\Preferences\Groups\Groups.xml") -Encoding utf8 -Append
"   <Group clsid=""{6D4A79E4-529C-4481-ABD0-F5BD7EA93BA7}"" name=""FSLogix Profile Include List"" image=""2"" changed=""$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")"" uid=""{6C2B4A05-A5F4-43E3-8F12-1C80538C2C6A}""><Properties action=""U"" newName="""" description=""Members of this group are on the include list for dynamic profiles"" deleteAllUsers=""1"" deleteAllGroups=""1"" removeAccounts=""0"" groupSid="""" groupName=""FSLogix Profile Include List""><Members><Member name=""$FSLogixProfileGroup"" action=""ADD"" sid=""""/></Members></Properties></Group>" | Out-File $("\\" + $env:USERDNSDOMAIN + "\SYSVOL\" + $env:USERDNSDOMAIN + "\Policies\{" + $GPOObject.GpoId + "}\Machine\Preferences\Groups\Groups.xml") -Encoding utf8 -Append

Write-Host "GPO Einstellungen wurden erfolgreich gesetzt."  -ForegroundColor Green
####################################################################################
Remove-Item "C:\Temp\fslogix\install\FSLogixAppsSetup.zip"

# Installieren + Neustart
# Window options
$windowInstallFsLogix.Owner = [System.Windows.Application]::Current.MainWindow
$result = $windowInstallFsLogix.ShowDialog()

# Auswertung der Abfrage
if ($result -eq $true) {
    Write-Output "Installation von FSLogix gestartet ..."
    #Start-Process "C:\Temp\fslogix\install\x64\Release\FSLogixAppsSetup.exe" -ArgumentList "/install /quiet" -Wait -Passthru | Out-Null
	$path = $PSScriptRoot + "\" + $script002Name
	if (!(Test-Path $path)) {
		Write-Host "Das Script '$script002Name' konnte nicht gefunden werden! Bitte starte das Script manuell, um fortzufahren."  -ForegroundColor Red
	}
	else {
		Write-Host "Scriptpfad: $path"
		Write-Host ""
		& $path
	}
} else {
    Write-Output "FSLogix wird nicht installiert."
	exit 0
}
