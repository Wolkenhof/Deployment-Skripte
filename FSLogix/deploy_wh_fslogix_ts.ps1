<#
############################## Wolkenhof ##############################
Purpose : FSLogix Installationsscript for Terminal Server
Created : 27.11.2024
Source  : https://github.com/Wolkenhof/Deployment-Skripte
Author  : jgu
Company : Wolkenhof GmbH
############################## Wolkenhof ##############################
#>

Write-Host "FSLogix Installationsscript (002) [Version 1.0]"
Write-Host "Copyright (c) 2024 Wolkenhof GmbH."
Write-Host ""

function InstallFSLogix() {
	$devices = Get-ADComputer -Filter "OperatingSystem -like '*Windows Server*'" | Select-Object Name | Out-GridView -Title "Auf welchem Server(n) soll FSLogix installiert werden?" -PassThru
	if($null -eq $devices) {
		Write-Host "Abbrechen..."
		exit 1
	}
	
	$devices | ForEach-Object {
		$computerName = $_.Name
		
		# Check if Device is DC
		$isDomainController = (Get-ADDomainController -Filter {Name -eq $computerName} -ErrorAction SilentlyContinue) -ne $null
		if ($isDomainController) {
			Write-Host "--> Installation gestartet auf Server $computerName (Domain Controller)" -ForegroundColor Yellow
			
			$session = New-PSSession -ComputerName $computerName
			Invoke-Command -Session $session -ScriptBlock { New-Item -ItemType Directory -Force -Path C:\Temp | Out-Null}
			Copy-Item -Path $FSLogixInstallPath\x64\Release\FSLogixAppsSetup.exe -Force -Destination 'C:\Temp' -ToSession $session
			Invoke-Command -Session $session -ScriptBlock {
				Set-ExecutionPolicy -ExecutionPolicy Bypass
				Write-Host "Starte installation ..."
				Start-Process -FilePath C:\Temp\FSLogixAppsSetup.exe -ArgumentList "/install /quiet /norestart /log C:\Temp\FSLogix_Log.txt" -Wait
				Write-Host "Installation beendet"
			}
			Copy-Item -Path "C:\Temp\FSLogix_Log.txt" -Force -Destination $FSLogixInstallPath\FSLogix_Log_$computerName.txt -FromSession $session

			Remove-PSSession -Session $session
			return
		}
		else {
			Write-Host "--> Installation gestartet auf Server $computerName (Terminalserver)" -ForegroundColor Yellow
			
			$session = New-PSSession -ComputerName $computerName
			Invoke-Command -Session $session -ScriptBlock { New-Item -ItemType Directory -Force -Path C:\Temp | Out-Null }
			Copy-Item -Path $FSLogixInstallPath\x64\Release\FSLogixAppsSetup.exe -Force -Destination 'C:\Temp' -ToSession $session
			Invoke-Command -Session $session -ScriptBlock {
				Set-ExecutionPolicy -ExecutionPolicy Bypass
				Write-Host "Starte installation ..."
				Start-Process -FilePath C:\Temp\FSLogixAppsSetup.exe -ArgumentList "/install /quiet /norestart /log C:\Temp\FSLogix_Log.txt" -Wait
				Write-Host "Installation beendet"
				
				try {
					Write-Host "Entferne 'Jeder' aus der Gruppe 'FSLogix ODFC Include List' ..."
					Remove-LocalGroupMember -Group "FSLogix ODFC Include List" -Member Jeder
				}
				catch {
					if ($_.Exception.Message -like "*MemberNotFound*") {
						Write-Warning "'Jeder' existiert bereits nicht in der Gruppe 'FSLogix ODFC Include List'."
					} else {
						Write-Error "Ein Fehler ist aufgetreten: $_.Exception.Message"
					}
				}

				try {
					Write-Host "Entferne 'Jeder' aus der Gruppe 'FSLogix Profile Include List' ..."
					Remove-LocalGroupMember -Group "FSLogix Profile Include List" -Member Jeder
				}
				catch {
					if ($_.Exception.Message -like "*MemberNotFound*") {
						Write-Warning "'Jeder' existiert bereits nicht in der Gruppe 'FSLogix Profile Include List'."
					} else {
						Write-Error "Ein Fehler ist aufgetreten: $_.Exception.Message"
					}
				}

				$localGroupProfile = "FSLogix Profile Include List"
				$localGroupO365 = "FSLogix ODFC Include List"
				$domainGroupProfile = "grp_Profile_Container"
				$domainGroupO365 = "grp_O365_Container"

				$group = [ADSI]"WinNT://./$localGroupProfile,group"
				$members = @($group.psbase.Invoke("Members")) | ForEach-Object { $_.GetType().InvokeMember("Name", 'GetProperty', $null, $_, $null) }

				try {
					if ($members -contains $localGroupProfile) {
						Write-Host "Die Gruppe '$localGroupProfile' ist bereits Mitglied der Gruppe '$domainGroupProfile'!" -ForegroundColor Yellow
					} else {
						$group.Add("WinNT://$domainGroupProfile,group")
						Write-Host "Die Gruppe '$localGroupProfile' wurde zur Gruppe '$domainGroupProfile' hinzugefügt." -ForegroundColor Green
					}
				}
				catch {
					if ($_.Exception.Message -like "*already a member*") {
						Write-Host "Die Gruppe '$localGroupProfile' ist bereits Mitglied der Gruppe '$domainGroupProfile'!" -ForegroundColor Yellow
					} else {
						Write-Error "Ein Fehler ist aufgetreten: $_.Exception.Message"
					}
				}

				$group = [ADSI]"WinNT://./$localGroupO365,group"
				$members = @($group.psbase.Invoke("Members")) | ForEach-Object { $_.GetType().InvokeMember("Name", 'GetProperty', $null, $_, $null) }

				try {
					if ($members -contains $localGroupO365) {
						Write-Host "Die Gruppe '$localGroupO365' ist bereits Mitglied der Gruppe '$domainGroupO365'!" -ForegroundColor Yellow
					} else {
						$group.Add("WinNT://$domainGroupO365,group")
						Write-Host "Die Gruppe '$localGroupO365' wurde zur Gruppe '$domainGroupO365' hinzugefügt." -ForegroundColor Green
					}
				}
				catch {
					if ($_.Exception.Message -like "*already a member*") {
						Write-Host "Die Gruppe '$localGroupO365' ist bereits Mitglied der Gruppe '$domainGroupO365'!" -ForegroundColor Yellow
					} else {
						Write-Error "Ein Fehler ist aufgetreten: $_.Exception.Message"
					}
				}

			}
			Copy-Item -Path "C:\Temp\FSLogix_Log.txt" -Force -Destination $FSLogixInstallPath\FSLogix_Log_$computerName.txt -FromSession $session
			
			Remove-PSSession -Session $session
			return
		}
	}
	Write-Host "FSLogix wurde auf allen ausgewählten Servern installiert."  -ForegroundColor Green
	exit 0
}

function UpdateFSLogix() {
	$devices = Get-ADComputer -Filter "OperatingSystem -like '*Windows Server*'" | Select-Object Name | Out-GridView -Title "Auf welchem Server(n) soll FSLogix installiert werden?" -PassThru
	if($null -eq $devices) {
		Write-Host "Abbrechen..."
		exit 1
	}

	$devices | ForEach-Object {
		$computerName = $_.Name
		Write-Host "--> Update gestartet auf Server $computerName" -ForegroundColor Yellow
		
		$session = New-PSSession -ComputerName $computerName
		Invoke-Command -Session $session -ScriptBlock {
			Add-Type -AssemblyName System.Net.Http

			function Get-FSLogixVersion {
				param (
					[string]$RegistryPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
				)
				
				$subkeys = Get-ChildItem -Path $RegistryPath
				
				foreach ($subkey in $subkeys) {
					$properties = Get-ItemProperty -Path $subkey.PSPath
					
					if ($properties.DisplayName -eq "Microsoft FSLogix Apps") {
						return $properties.DisplayVersion
					}
				}
				
				Write-Warning "FSLogix wurde nicht gefunden."
				return $null
			}
			function Show-Progress {
				param (
					[int]$PercentComplete
				)
				Write-Progress -Activity "Downloading FSLogix" -Status "$PercentComplete% Complete" -PercentComplete $PercentComplete
			}
			
			Set-ExecutionPolicy -ExecutionPolicy Bypass
			New-Item -ItemType Directory -Force -Path C:\Temp | Out-Null
			
			$FslogixUrl= "https://aka.ms/fslogix_download"
			$value = Get-FSLogixVersion

			# Neuste Version
			$request = [System.Net.HttpWebRequest]::Create($FslogixUrl)
			$request.Method = "HEAD"
			$request.AllowAutoRedirect = $true
			$response = $request.GetResponse()
			$finalUrl = $response.ResponseUri.AbsoluteUri
			$fileName = [System.IO.Path]::GetFileName($finalUrl)
			$regexPattern = "\d+\.\d+\.\d+\.\d+"
			$neueVersion = 0.0
			
			# Neuste Version
			try {
				Write-Host "Suche nach neue FSLogix Version ..."
				$request = [System.Net.HttpWebRequest]::Create($FslogixUrl)
				$request.Method = "HEAD"
				$request.AllowAutoRedirect = $true
				$request.UserAgent = "Wolkenhof\1.0 PowerShell FSLogixScript"
				$response = $request.GetResponse()
				$finalUrl = $response.ResponseUri.AbsoluteUri
				$fileName = [System.IO.Path]::GetFileName($finalUrl)
				$regexPattern = "\d+\.\d+\.\d+\.\d+"
				$neueVersion = 0.0

				if ($fileName -match $regexPattern) {
					$neueVersion = $matches[0]
				}

				Write-Host Aktuelle Version: $value
				Write-Host Neue Version: $neueVersion
				Write-Host ""
			}
			catch {
				Write-Error "Fehler bei der Webanfrage: $_.Exception.Message"
				exit 1
			}

			if ($value -eq $neueVersion) {
				Write-Host "FSLogix ist aktuell. Überspringe Download ..."
			}
			else {
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
				
				Write-Host "Starte Installation ..."
				Start-Process -FilePath C:\Temp\fslogix\install\x64\Release\FSLogixAppsSetup.exe -ArgumentList "/install /quiet /norestart" -Wait
				Write-Host "Installation beendet"
			}
		}

		Remove-PSSession -Session $session
	}
	Write-Host "FSLogix wurde auf allen ausgewählten Servern aktualisiert."  -ForegroundColor Green
	exit 0
}

Import-Module ActiveDirectory
Add-Type -AssemblyName PresentationFramework
$DCHostname = Get-ADDomainController -Filter * | Select-Object -ExpandProperty HostName
$FSLogixInstallPath = "\\$DCHostname\FXLogix_Install$"
$global:Action = "none"

### WPF Logic ###
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="FSLogix Server-Installation | Wolkenhof GmbH" Height="165" Width="460" ResizeMode="NoResize">
    <Grid>
        <Label Content="Wie lautet der UNC-Pfad zur FSLogix-Installation Freigabe?" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top"/>
        <TextBox x:Name="TextBox1" HorizontalAlignment="Left" Margin="10,40,0,0" VerticalAlignment="Top" Width="425" Text="$FSLogixInstallPath"/>
        <Button Content="OK" x:Name="OK_Button" HorizontalAlignment="Left" Margin="360,97,0,0" VerticalAlignment="Top" Width="75"/>
        <Button Content="Abbrechen" x:Name="Cancel_Button" HorizontalAlignment="Left" Margin="275,97,0,0" VerticalAlignment="Top" Width="75"/>
    </Grid>
</Window>
"@
$reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]$xaml)
$window = [System.Windows.Markup.XamlReader]::Load($reader)

# Button Events
$okButton = $window.FindName("OK_Button")
$cancelButton = $window.FindName("Cancel_Button")
$textBox1 = $window.FindName("TextBox1")

$okButton.Add_Click({
    $window.Tag = @{
        TextBox1Value = $textBox1.Text
    }
    $window.DialogResult = $true
    $window.Close()
})

$cancelButton.Add_Click({
    $window.DialogResult = $false
    $window.Close()
})

### Main Menu ###
$xamlMainMenu = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="FSLogix Server-Installation | Wolkenhof GmbH" Height="150" Width="460" ResizeMode="NoResize">
    <StackPanel Margin="5">
        <Label Content="Welche Aktion soll durchgeführt werden?"/>
        <RadioButton x:Name="ServerInstall" Content="FSLogix auf (mehreren) Server installieren" IsChecked="true" Margin="10,10,0,0"/>
        <RadioButton x:Name="ServerUpdate" Content="FSLogix aktualisieren" Margin="10,10,0,0" />
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
            <Button Content="OK" x:Name="OK_Button" Width="75"/>
            <Button Content="Abbrechen" x:Name="Cancel_Button" Width="75" Margin="10,0,0,0"/>
        </StackPanel>
    </StackPanel>
</Window>
"@
$readerMainMenu = [System.Xml.XmlReader]::Create([System.IO.StringReader]$xamlMainMenu)
$windowMainMenu = [System.Windows.Markup.XamlReader]::Load($readerMainMenu)

# Button Events
$okButtonMainMenu = $windowMainMenu.FindName("OK_Button")
$cancelButtonMainMenu = $windowMainMenu.FindName("Cancel_Button")
$serverInstallMainMenu = $windowMainMenu.FindName("ServerInstall")
$serverUpdateMainMenu = $windowMainMenu.FindName("ServerUpdate")

$okButtonMainMenu.Add_Click({
    if ($serverInstallMainMenu.IsChecked) {  
		$global:Action = "Install"
	}
	if ($serverUpdateMainMenu.IsChecked) {  
		$global:Action = "Update"
	}
	$windowMainMenu.DialogResult = $true
    $windowMainMenu.Close()
})

$cancelButtonMainMenu.Add_Click({
    $windowMainMenu.DialogResult = $false
    $windowMainMenu.Close()
	exit 1
})

##################################################################################
#$window.Owner = [System.Windows.Application]::Current.MainWindow
#$result = $window.ShowDialog()
#if ($result -eq $true) {
#    $FSLogixInstallPath = $window.FindName("TextBox1").Text
#} else {
#    Write-Output "Der Vorgang wurde abgebrochen."
#	exit 1
#}

$windowMainMenu.Owner = [System.Windows.Application]::Current.MainWindow
$resultMainMenu = $windowMainMenu.ShowDialog()

if ($global:Action -eq "Install") {
	Write-Host "Aktion: Installiere FSLogix"
	InstallFSLogix
}
if ($global:Action -eq "Update") {
	Write-Host "Aktion: Aktualisiere FSLogix"
	UpdateFSLogix
}
if ($resultMainMenu -eq $false) {
	Write-Host "Der Vorgang wurde abgebrochen."
	exit 1
}
Write-Host "Unbekannter Fehler!"
exit 1
