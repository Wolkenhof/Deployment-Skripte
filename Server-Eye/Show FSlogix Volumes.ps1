<#
############################## Wolkenhof ##############################
Purpose : FSLogix User Disk Prüfer für Server-Eye
Created : 27.11.2024
Source  : https://github.com/Wolkenhof/Deployment-Skripte
Author  : jgu
Company : Wolkenhof GmbH
############################## Wolkenhof ##############################
#>

[CmdletBinding()]

Param
(
    [string]$minPercentage,
    [double]$minGB
)

Write-Host "FSLogix User Disk Pruefer [Version 2.1]"
Write-Host "Copyright (c) 2024 Wolkenhof GmbH"
Write-Host ""
Write-Host "Hinweis: Zum Anpassen des Schwellenwertes, gehe zu 'Sensor-Einstellungen -> Argumente'"
Write-Host "Hinweis: und gebe den gewuenschten Wert ein."
Write-Host ""
Write-Host "Hinweis: -minPercentage   =   minimal freier Speicherplatz (in %)"
Write-Host "Hinweis: -minGB           =   minimal freier Speicherplatz (in GB)"
Write-Verbose "Found $($partitions.Count) virtual disk partitions"

[array]$partitions = @( Get-Partition | Where-Object { $_.DiskId -match '&ven_msft&prod_virtual_disk' -and ! $_.DriveLetter -and (($_.Type -eq 'Basic') -or ($_.Type -eq 'IFS')) } )
if( ! $partitions -or ! $partitions.Count )
{
    Write-Host ""
    Write-Host "Kein Benutzer ist zur Zeit angemeldet."
    exit 0
}

[array]$fixedVolumes = @( Get-Volume | Where-Object { $_.DriveType -eq 'Fixed' } )
if( ! $fixedVolumes -or ! $fixedVolumes.Count )
{
    Write-Warning "Unable to find any fixed volumes"
}
else
{
    Write-Verbose "Found $($fixedVolumes.Count) fixed volumes"
}

[array]$virtualDisks = @( Get-Disk | Where-Object { $_.BusType -eq 'File Backed Virtual' } )
if( ! $virtualDisks -or ! $virtualDisks.Count )
{
    Write-Warning "Unable to find any file backed virtual disks"
}
else
{
    Write-Verbose "Found $($virtualDisks.Count) file backed virtual disks"
}

[int]$counter = 0
[int]$failCounter = 0;

[array]$results = @( ForEach( $partition in $partitions )
{
    $counter++
    $partitionGuid = $partition.Guid
    if ([string]::IsNullOrEmpty($partitionGuid))
    {
        Write-Verbose "Partition does not have an GUID, using AccessPaths instead ..."
        $partitionGuid = [regex]::Match($partition.AccessPaths, '\{[0-9a-fA-F-]{36}\}').Value
    }

    Write-Verbose "$counter / $($partitions.Count) : Partition GUID $partitionGuid"

    $volume = $fixedVolumes | Where-Object { $_.UniqueId -match $partitionGuid }
    if( !$volume )
    {
        Write-Warning "Unable to find fixed volume with GUID $partitionGuid"
        return
    }
    
    [string]$uniqueId = ($partition.UniqueId -split '[{}]')[-1]
    $disk = $virtualDisks | Where-Object { $_.UniqueId -eq $uniqueId }
    if( ! $disk )
    {
        Write-Warning "Unable to find disk with unique id $uniqueId"
        return
    }

    $fslogixRegValue = Get-ItemProperty -Path "HKLM:\SOFTWARE\FSLogix\Profiles\Sessions\*" -ErrorAction SilentlyContinue | Where-Object { $_.Volume -eq $volume.Path }
    if( ! $fslogixRegValue )
    {
        Write-Verbose "Registry-Eintrag fuer Volume $($volume.Path) nicht gefunden, ueberspringe..."
        continue
    }
    
    $profilePath = $fslogixRegValue.ProfilePath
    $volumeSize = $volume.Size / 1GB
    $sizeRemaining = $volume.SizeRemaining / 1GB
    $usedSize = $volumeSize - $sizeRemaining
    $usedPercentage = ($usedSize / $volumeSize) * 100
    $freePercentage = 100 - $usedPercentage

    $roundPercentage = [math]::Round($usedPercentage, 1)
    $roundUsedSize = [math]::Round($usedSize, 2)
    $roundVolumeSize = [math]::Round($volumeSize, 2)
    $roundFreePercentage = [math]::Round($freePercentage, 1)
    $roundSizeRemaining = [math]::Round($sizeRemaining, 2)

    $zustand = "OK"
    if ($freePercentage -lt $minPercentage) {
        $zustand = "Schwellwert ueberschritten"
        $failCounter++
    }

    if ($minGB -and $usedSize -gt $minGB) {
        $zustand = "Schwellwert ueberschritten"
        $failCounter++
    }

    $result = [pscustomobject][ordered]@{
        'Profil' = $profilePath
        'FreierSpeicher' = "$roundFreePercentage% ($roundSizeRemaining / $roundVolumeSize GBytes)"
        'Zustand' = $zustand
    }
    $result
})

$results | Format-Table @{
    Label="Profil";
    Expression={$_.Profil};
    Width=35
}, @{
    Label="Freier Speicher";
    Expression={$_.FreierSpeicher};
    Width=35
}, @{
    Label="Zustand";
    Expression={
        switch ($_.Zustand) {
            {$_ -eq "OK"} {
                $color = "$([char]27)[0;32m" # Grün
            }
            {$_ -ne "OK" } {
                $color = "$([char]27)[0;31m" # Rot
            }
        }
        $_.Zustand
        };
    Width=30
}

if ($failCounter -gt 0) {
    exit 1
}
exit 0
