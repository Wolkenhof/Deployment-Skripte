
<#
.SYNOPSIS
This script is a wrapper for the OfficeDeploymentTool to use parameters instead of a predefined XML file.

.PARAMETER Info
Show information on installed Office365 products

.PARAMETER Update
Update installed Office365 products

.PARAMETER Install
Install Office365 product
Possible values are given as a comma-seperated list of these elements O365ProPlusRetail, O365BusinessRetail, VisioProRetail, ProjectProRetail

.PARAMETER Uninstall
Uninstall Office365 product
Possible values are given as a comma-seperated list of these elements O365ProPlusRetail, O365BusinessRetail, VisioProRetail, ProjectProRetail, All

.PARAMETER Online
Use Online CDN instead of ISO-Storage as source for install or update

.PARAMETER ExcludeAppID
Choose which apps should be excluded from install
Possible values are given as a comma-seperated list of these elements Access, Bing, Excel, Groove, Lync, OneDrive, OneNote, Outlook, PowerPoint, Publisher, Teams, Word

.PARAMETER LanguageID
Choose which languages should be installed od uninstalled
culture codes are given as a comma-seperated list e.g. de-de,en-us

.PARAMETER ClientEdition
Choose bitness of installation
Possible values are 32 or 64

.PARAMETER Channel
Chosse update channel for install or update
Possible values are Current or MonthlyEnterprise

.PARAMETER Version
Choose specific version to install or update

.PARAMETER SharedComputerLicensing
Activate SharedComputerLicensing for multiuser installation
#>

[CmdletBinding(DefaultParameterSetName='Update')]

param (
	[Parameter(ParameterSetName='Info')]
	[switch]$Info,
	
	[Parameter(ParameterSetName='Update')]
	[switch]$Update = $true,

	[Parameter(ParameterSetName='Install', Mandatory)]
	[ValidateSet('O365ProPlusRetail','O365BusinessRetail','VisioProRetail','ProjectProRetail')]
	[string[]]$Install,

	[Parameter(ParameterSetName='Uninstall', Mandatory)]
	[ValidateSet('O365ProPlusRetail','O365BusinessRetail','VisioProRetail','ProjectProRetail', 'All')]
	[string[]]$Uninstall,

	[Parameter(ParameterSetName='Update')]
	[Parameter(ParameterSetName='Install')]
	[switch]$Online,

	[Parameter(ParameterSetName='Install')]
	[ValidateSet('Access','Bing','Excel','Groove','Lync','OneDrive','OneNote','Outlook','PowerPoint','Publisher','Teams','Word')]
	[string[]]$ExcludeAppID = @('Bing', 'Groove', 'Lync', 'OneDrive', 'Teams'),

	[Parameter(ParameterSetName='Install')]
	[Parameter(ParameterSetName='Uninstall')]
	[ValidatePattern('^[a-z,A-Z]{2}-[a-z,A-Z]{2}?$')]
	[string[]]$LanguageID = "de-de",

	[Parameter(ParameterSetName='Install')]
	[ValidateSet('32','64')]
	[String]$ClientEdition,

	[Parameter(ParameterSetName='Install')]
	[Parameter(ParameterSetName='Update')]
	[ValidateSet('Current','MonthlyEnterprise')]
	[String]$Channel,

	[Parameter(ParameterSetName='Install')]
	[Parameter(ParameterSetName='Update')]
	[ValidatePattern('^16.0.[0-9]{5}.[0-9]{5}?$')]
	[String]$Version,

	[Parameter(ParameterSetName='Install')]
	[ValidateSet('0','1')]
	[String]$SharedComputerLicensing
)

Begin
{
	$Verbose = ('-Verbose' -in $MyInvocation.UnboundArguments -or $MyInvocation.BoundParameters.ContainsKey('Verbose'))

	Write-Output "Using ParameterSet $($PsCmdlet.ParameterSetName)"
	Write-Verbose "    Online: $Online"
	Write-Verbose "    Install: $Install"
	Write-Verbose "    Uninstall: $Uninstall"
	Write-Verbose "    LanguageID: $LanguageID"
	Write-Verbose "    ExcludeAppID: $ExcludeAppID"
	Write-Verbose "    ClientEdition: $ClientEdition"
	Write-Verbose "    Channel: $Channel"
	Write-Verbose "    Version: $Version"
	Write-Verbose "    SharedComputerLicensing: $SharedComputerLicensing"
	
	$M365Path = "C:\Wolkenhof\Office\Tools"
	$ODT = "$M365Path\Office365\Setup\setup.exe"
    Write-Verbose "Local Office: $M365Path"
	Write-Verbose "ODT: $ODT"

	If (-not (Test-Path -Path $ODT))
	{
		Write-Error "Could not detect ODT!"
		Exit 1  
	}

	$TempConfigXML="$env:Temp\TempConfig.xml"
	$C2RConfigurationPath = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
	$C2RInventoryPath = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Inventory\Office\16.0"

	Write-Output ""
	Write-Output "Detecting installed Office"

	If ($O365Installed = (Test-Path $C2RConfigurationPath) -and (Test-Path $C2RInventoryPath))
	{
		Write-Output "O365 installed: $O365Installed"

		$Platform=(Get-ItemProperty -Path $C2RConfigurationPath).Platform
		Write-Verbose "    Platform: $Platform"

		$AudienceData=((Get-ItemProperty -Path $C2RConfigurationPath).AudienceData).Split(':')[2]
		Write-Verbose "    AudienceData: $AudienceData"
		
		$UpdateUrl = (Get-ItemProperty -Path $C2RConfigurationPath).UpdateUrl
		Write-Verbose "    UpdateUrl: $UpdateUrl"
		
		$UpdateChannel = (Get-ItemProperty -Path $C2RConfigurationPath).UpdateChannel
		Write-Verbose "    UpdateChannel: $UpdateChannel"
		
		$CDNBaseUrl = (Get-ItemProperty -Path $C2RConfigurationPath).CDNBaseUrl
		Write-Verbose "    CDNBaseUrl: $CDNBaseUrl"
		
		If (-not $SharedComputerLicensing)
		{
			$SharedComputerLicensing = (Get-ItemProperty -Path $C2RConfigurationPath).SharedComputerLicensing
			Write-Verbose "    SharedComputerLicensing: $SharedComputerLicensing"
		}
		
		$UpdatesEnabled = (Get-ItemProperty -Path $C2RConfigurationPath).UpdatesEnabled
		Write-Verbose "    UpdatesEnabled: $UpdatesEnabled"
		
		$VersionToReport = (Get-ItemProperty -Path $C2RConfigurationPath).VersionToReport
		Write-Verbose "    VersionToReport: $VersionToReport"
		
		$ProductReleaseIds = (Get-ItemProperty -Path $C2RConfigurationPath).ProductReleaseIds -split ','
		Write-Verbose "    ProductReleaseIds: $ProductReleaseIds"
		
		$ProductReleaseIdsValues = @{}
		ForEach ($id in $ProductReleaseIds)
		{
			Write-Verbose "    ProductReleaseId: $id"
			$ReleaseValues = @{}
			$items = (Get-Item -Path $C2RConfigurationPath).Property | Where {$_ -like "$id*"}
			foreach ($item in $items)
			{
				$value = $(Get-ItemProperty -Path $C2RConfigurationPath).$item
				$ReleaseValues[$item.split(".")[1]] = $value
				Write-Verbose "        $($item.split(".")[1]): $value" 
			}
			$ProductReleaseIdsValues[$id] = $ReleaseValues
		}
		
		$BingAddon = (Get-ItemProperty -Path $C2RConfigurationPath).BingAddon
		Write-Verbose "    BingAddon: $BingAddon"
		
		$TeamsAddon = (Get-ItemProperty -Path $C2RConfigurationPath).TeamsAddon
		Write-Verbose "    TeamsAddon: $TeamsAddon"
		
		$OneDriveClientAddon = (Get-ItemProperty -Path $C2RConfigurationPath).OneDriveClientAddon
		Write-Verbose "    OneDriveClientAddon: $OneDriveClientAddon"

		$OfficeProductReleaseIds = ((Get-ItemProperty -Path $C2RInventoryPath).OfficeProductReleaseIds).Split(',')
		Write-Verbose "    OfficeProductReleaseIds: $OfficeProductReleaseIds"
		
		$OfficeCultures = ((Get-ItemProperty -Path $C2RInventoryPath).OfficeCulture).Split(',')
		Write-Verbose "    OfficeCultures: $OfficeCultures"
		
		$OfficePackageVersion = (Get-ItemProperty -Path $C2RInventoryPath).OfficePackageVersion
		Write-Verbose "    OfficePackageVersion: $OfficePackageVersion"
	}
}

Process
{
	Switch ($PsCmdlet.ParameterSetName)
	{
		'Info'
		{
			Write-Output "O365 installed: $O365Installed"
			Write-Output "    ProductReleaseIds: $ProductReleaseIds"
			Write-Output "    Version: $VersionToReport"
			Write-Output "    Platform: $Platform"
			If ($AudienceData -eq "MEC" )
			{
	            Write-Output "    UpdateChannel: MonthlyEnterprise"
			}
			Else
			{
	            Write-Output "    UpdateChannel: Current"
			}
			Write-Output "    UpdatesEnabled: $UpdatesEnabled"
			Write-Output "    SharedComputerLicensing: $SharedComputerLicensing"
			break
		}
		'Update'
		{
			Write-Output ""
			Write-Output "Building ConfigXML for Update"
			$temp = '<Configuration>'
			Write-Verbose $temp
			Set-Content -Path $TempConfigXML -Value $temp
			
			# parameter wins over reg
			If (-not ($Channel))
			{
				# makes Current the default, if not specified
				If ($AudienceData -eq "MEC" )
				{
	                $Channel="MonthlyEnterprise"
				}
				Else
				{
	                $Channel="Current"
				}
			}
			
			If ($Platform -eq "x64")
			{
				$ClientEdition = "64"
			}
			Else
			{
				$ClientEdition = "32"
			}
			
			If ($Version)
			{
				If (-not (Test-Path -Path $($M365Path + '\Office365\Source\' + $Channel + '\Office\Data\' + $Version)))
				{
					Write-Error "Source for Version $Version not found!"
					Exit 1
				}
				$tempVersion = ' Version="' + $Version + '"'  
			}
			Else
			{
				$tempVersion =''
			}

			If (-not $online)
			{
				$SourcePath = "$M365Path\Office365\Source\$Channel"
				$temp = '  <Add SourcePath="' + $SourcePath + '" OfficeClientEdition="' + $ClientEdition + '" Channel="' + $Channel + '"' + $tempVersion + ' AllowCdnFallback="True">'
			}
			Else
			{
				$temp = '  <Add OfficeClientEdition="' + $ClientEdition + '" Channel="' + $Channel + '"' + $tempVersion + ' AllowCdnFallback="True">'
			}
			Write-Verbose $temp
			Add-Content -Path $TempConfigXML -Value $temp
			
			ForEach ($Id In $OfficeProductReleaseIds)
			{
				$temp = '    <Product ID="' + $Id + '" >'
				Write-Verbose $temp
	            Add-Content -Path $TempConfigXML -Value $temp

				ForEach ($Oc In $OfficeCultures)
				{
		            $temp = '      <Language ID="'+ $Oc + '" />'
					Write-Verbose $temp
		            Add-Content -Path $TempConfigXML -Value $temp
	            }

				ForEach ($Ea In $($ProductReleaseIdsValues[$id].ExcludedApps).Split(','))
				{
		            $temp = '      <ExcludeApp ID="' + $Ea + '" />'
					Write-Verbose $temp
		            Add-Content -Path $TempConfigXML -Value $temp
				}

				$temp = '    </Product>'
				Write-Verbose $temp
	            Add-Content -Path $TempConfigXML -Value $temp
			}
			
			$temp = '  </Add>'
			Write-Verbose $temp
			Add-Content -Path $TempConfigXML -Value $temp
			
			If ($SharedComputerLicensing)
			{
	            $temp = '  <Property Name="SharedComputerLicensing" Value="' + $SharedComputerLicensing + '" />'
				Write-Verbose $temp
	            Add-Content -Path $TempConfigXML -Value $temp
			}
			
			$temp = '  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />'
			Write-Verbose $temp
			Add-Content -Path $TempConfigXML -Value $temp
			
			If (-not $online)
			{
				$UpdatePath = "$M365Path\Office365\Source\$Channel"
				$temp = '  <Updates Enabled="TRUE" UpdatePath="' + $UpdatePath + '" />'
			}
			Else
			{
				$temp = '  <Updates Enabled="TRUE" />'
			}
			Write-Verbose $temp
			Add-Content -Path $TempConfigXML -Value $temp
			
			$temp = '  <Display Level="None" AcceptEULA="TRUE" />'
			Write-Verbose $temp
			Add-Content -Path $TempConfigXML -Value $temp
			
			$temp = '  <Logging Name="Office365Setup.txt" Level="Standard" Path="C:\Wolkenhof\Logs\Office365.log" />'
			Write-Verbose $temp
			Add-Content -Path $TempConfigXML -Value $temp
			
			$temp = '</Configuration>'
			Write-Verbose $temp
			Add-Content -Path $TempConfigXML -Value $temp
			Write-Output ""
			Write-Output "Running OfficeDeploymentTool"
			Write-Verbose "&$ODT /configure $TempConfigXML"
			&$ODT /configure $TempConfigXML
			break
		}
		'Install'
		{
			Write-Output ""
			Write-Output "Building ConfigXML for Install"
			$temp = '<Configuration>'
			Write-Verbose $temp
			Set-Content -Path $TempConfigXML -Value $temp
			
			# parameter wins over reg
			If (-not ($Channel))
			{
				# makes Current the default, if not specified
				If ($AudienceData -eq "MEC" )
				{
	                $Channel="MonthlyEnterprise"
				}
				Else
				{
	                $Channel="Current"
				}
			}

			# reg wins over parameter
			# make 64Bit the default if not specified or installed
			If ($ClientEdition -eq "32")
			{
				If ($Platform -eq "x64")
				{
					$ClientEdition = "64"
				}
			}
			Else
			{
				$ClientEdition = "64"
			}

			If ($Version)
			{
				If ((-not $online) -And (-not (Test-Path -Path $($M365Path + '\Office365\Source\' + $Channel + '\Office\Data\' + $Version))))
				{
					Write-Error "Source for Version $Version not found!"
					Exit 1
				}
				$tempVersion = ' Version="' + $Version + '"'  
			}
			Else
			{
				$tempVersion =''
			}
			
			If (-not $online)
			{
				$SourcePath = "$M365Path\Office365\Source\$Channel"
				$temp = '  <Add SourcePath="' + $SourcePath + '" OfficeClientEdition="' + $ClientEdition + '" Channel="' + $Channel + '"' + $tempVersion + ' AllowCdnFallback="True">'
			}
			Else
			{
				$temp = '  <Add OfficeClientEdition="' + $ClientEdition + '" Channel="' + $Channel + '"' + $tempVersion + ' AllowCdnFallback="True">'
			}
			Write-Verbose $temp
			Add-Content -Path $TempConfigXML -Value $temp
			
			#merge ExcludeApp
			ForEach ($Id In $OfficeProductReleaseIds)
			{
				$ExcludeAppID = ($ExcludeAppID + $(($ProductReleaseIdsValues[$id].ExcludedApps).Split(',')) ) | Sort-Object -Unique
			}
			
			# merge Language
			$OfficeCultures = ($OfficeCultures + $LanguageID) | Sort-Object -Unique

			# merge Product
			$OfficeProductReleaseIds = ($OfficeProductReleaseIds + $Install) | Sort-Object -Unique        
			ForEach ($Id In $OfficeProductReleaseIds)
			{
				$temp = '    <Product ID="' + $Id + '" >'
				Write-Verbose $temp
	            Add-Content -Path $TempConfigXML -Value $temp
				
				ForEach ($Oc In $OfficeCultures)
				{
		            $temp = '      <Language ID="'+ $Oc + '" />'
					Write-Verbose $temp
		            Add-Content -Path $TempConfigXML -Value $temp
	            }
	            
				ForEach ($Ea In $ExcludeAppID)
				{
		            $temp = '      <ExcludeApp ID="' + $Ea + '" />'
					Write-Verbose $temp
		            Add-Content -Path $TempConfigXML -Value $temp
				}
				$temp = '    </Product>'
				Write-Verbose $temp
	            Add-Content -Path $TempConfigXML -Value $temp
			}

			$temp = '  </Add>'
			Write-Verbose $temp
			Add-Content -Path $TempConfigXML -Value $temp
			
			If ($SharedComputerLicensing)
			{
	            $temp = '  <Property Name="SharedComputerLicensing" Value="' + $SharedComputerLicensing + '" />'
				Write-Verbose $temp
	            Add-Content -Path $TempConfigXML -Value $temp
			}
			
			$temp = '  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />'
			Write-Verbose $temp
			Add-Content -Path $TempConfigXML -Value $temp
			
			If (-not $online)
			{
				$UpdatePath = "$M365Path\Office365\Source\$Channel"
				$temp = '  <Updates Enabled="TRUE" UpdatePath="' + $UpdatePath + '" />'
			}
			Else
			{
				$temp = '  <Updates Enabled="TRUE" />'
			}
			Write-Verbose $temp
			Add-Content -Path $TempConfigXML -Value $temp
			
			$temp = '  <Display Level="None" AcceptEULA="TRUE" />'
			Write-Verbose $temp
			Add-Content -Path $TempConfigXML -Value $temp
			
			$temp = '  <Logging Name="Office365Setup.txt" Level="Standard" Path="' + $env:windir + '\System32\LogFiles\Office365" />'
			Write-Verbose $temp
			Add-Content -Path $TempConfigXML -Value $temp
			
			$temp = '</Configuration>'
			Write-Verbose $temp
			Add-Content -Path $TempConfigXML -Value $temp
			
			Write-Output ""
			Write-Output "Running OfficeDeploymentTool"
			Write-Verbose "&$ODT /configure $TempConfigXML"
			&$ODT /configure $TempConfigXML
			break
		}
		'Uninstall'
		{
			Write-Output ""
			Write-Output "Building ConfigXML for Uninstall"
			$temp = '<Configuration>'
			Write-Verbose $temp
			Set-Content -Path $TempConfigXML -Value $temp

			If ($Uninstall -contains "All")
			{
				$temp = '  <Remove All="TRUE">'
				Write-Verbose $temp
				Add-Content -Path $TempConfigXML -Value $temp
			}
			Else
			{
				$temp = '  <Remove>'
				Write-Verbose $temp
				Add-Content -Path $TempConfigXML -Value $temp
			
				ForEach ($Id In $Uninstall)
				{
					$temp = '    <Product ID="' + $Id + '" >'
					Write-Verbose $temp
	                Add-Content -Path $TempConfigXML -Value $temp
				
					ForEach ($La In $LanguageID)
					{
		                $temp = '      <Language ID="'+ $La + '" />'
						Write-Verbose $temp
		                Add-Content -Path $TempConfigXML -Value $temp
	                }

				    $temp = '    </Product>'
					Write-Verbose $temp
	                Add-Content -Path $TempConfigXML -Value $temp
				}
			}

			$temp = '  </Remove>'
			Write-Verbose $temp
			Add-Content -Path $TempConfigXML -Value $temp
			
			$temp = '  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />'
			Write-Verbose $temp
			Add-Content -Path $TempConfigXML -Value $temp
			
			$temp = '  <Display Level="None" AcceptEULA="TRUE" />'
			Write-Verbose $temp
			Add-Content -Path $TempConfigXML -Value $temp
			
			$temp = '  <Logging Name="Office365Setup.txt" Level="Standard" Path="' + $env:windir + '\System32\LogFiles\Office365" />'
			Write-Verbose $temp
			Add-Content -Path $TempConfigXML -Value $temp
			
			$temp = '</Configuration>'
			Write-Verbose $temp
			Add-Content -Path $TempConfigXML -Value $temp
			
			Write-Output ""
			Write-Output "Running OfficeDeploymentTool"
			Write-Verbose "&$ODT /configure $TempConfigXML"
			&$ODT /configure $TempConfigXML
			break
		}
	}	
}

End
{
    # clean up
}
