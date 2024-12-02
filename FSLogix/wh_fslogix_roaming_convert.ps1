##
## FSLogix Installationsscript (002) - Version 1.0
##

Write-Host "FSLogix Konvertierungsscript [Version 1.0]"
Write-Host "Copyright (c) 2024 Wolkenhof GmbH."
Write-Host ""

$newprofilepath = "\\DC01\FSLogixStore$"
Write-Host "FSLogix Freigabe: $newprofilepath"

$ENV:PATH="$ENV:PATH;C:\Program Files\fslogix\apps\"
$oldprofiles = gci c:\users | ?{$_.psiscontainer -eq $true} | select -Expand fullname | sort | out-gridview -OutputMode Multiple -title "WÃ¤hle aus, welche Profile konvertiert werden sollen."

# foreach old profile
foreach ($old in $oldprofiles) {
	$sam = ($old | split-path -leaf)
	$sid = (New-Object System.Security.Principal.NTAccount($sam)).translate([System.Security.Principal.SecurityIdentifier]).Value

	# set the nfolder path to \\newprofilepath\username_sid
	$nfolder = join-path $newprofilepath ($sid+"_"+$sam)
	# if $nfolder doesn't exist - create it with permissions
	if (!(test-path $nfolder)) {New-Item -Path $nfolder -ItemType directory | Out-Null}
	& icacls $nfolder /setowner "$env:userdomain\$sam" /T /C
	& icacls $nfolder /grant $env:userdomain\$sam`:`(OI`)`(CI`)F /T

	$vhd = Join-Path $nfolder ("Profile_"+$sam+".vhdx")

	frx.exe copy-profile -filename $vhd -sid $sid -verbose -size-mbs 50000 -dynamic 1
}
Write-Host "Alle Benutzer wurden konvertiert." -ForegroundColor Green 
