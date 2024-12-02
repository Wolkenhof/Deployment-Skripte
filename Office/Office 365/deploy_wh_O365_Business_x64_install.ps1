<#
############################## Wolkenhof ##############################
Purpose : Install or Update O365 Apps for Business x64
Created : 27.11.2024
Source  : https://github.com/Wolkenhof/Deployment-Skripte
Author  : jgu
Company : Wolkenhof GmbH
############################## Wolkenhof ##############################
#>

powershell -ep Bypass "C:\Wolkenhof\Office\Tools\Office365\Deploy-Office365.ps1" -Install O365BusinessRetail -ClientEdition 64 -Channel Current -Online -LanguageID de-de -SharedComputerLicensing 1 -ExcludeAppID Groove,Lync,OneDrive,Teams -Verbose
