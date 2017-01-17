<#

.SYNOPSIS

Performs the following actions:

1) Confirms that only one AD and one OrgId setting exists for all app MRUs and for roaming settings

2) If more than one AD exists, this script does nothing.
   If more than one AAD exists, this script does nothing.
   If no AD exists, this script does nothing.
   If no AAD exists (i.e the user hasn't been migrated to AAD, OR, the user hasn't launched the given Office app before), this script does nothing.

3) Takes a backup of the OrgId settings in all app MRus and roaming, writing it to the path provided in the "-BackupPath" argument

4) Makes the following changes:

    Most Recently Used (MRU)
    ------------------------
    Moves FROM:
	    HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\15.0\**App**\User MRU\AD_**\
    Moves TO:
	    HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\15.0\**App**\User MRU\OrgId_**\

    Roaming settings and customizations
    -----------------------------------
    Moves FROM:
	    HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\15.0\Common\Roaming\Identities\**_AD
    Moves TO:
	    HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\15.0\Common\Roaming\Identities\**_OrgId

5) If the "-Undo" argument is specified, applies the backup registry files located in the specified "-BackupPath" argument

#>
param(
   [parameter(Mandatory=$true)]
   [string]$BackupPath,
   [switch]$Undo
)

$basePath = "HKCU:\Software\Microsoft\Office\15.0\"
$basePathForReg = "HKCU\Software\Microsoft\Office\15.0\"


## Undo if undo switch specified

if ($Undo)
{
    ## Ensure backup files exist
    if ((Test-Path "$BackupPath\OfficeSettingsMigrateBackup_orgIdUserRoaming.reg") -eq $false)
    {
        Write-Host "No backup files exist"
        return
    }

    ## Delete existing OrgId MRU and Roaming
    Get-ChildItem "$basePath\*\User MRU\OrgId_*" | Remove-Item -Recurse
    Get-ChildItem "$basePath\Common\Roaming\Identities\*_OrgId" | Remove-Item -Recurse

    foreach ($backupFile in Get-ChildItem "$BackupPath\OfficeSettingsMigrateBackup_*")
    {
        Write-Host "Restoring from $backupFile"
        reg.exe import $backupFile
    }

    return
}


## Get AD and OrgId key names

$mruAdUserIds = Get-ChildItem "$basePath\*\User MRU\AD_*" | 
  foreach { ($_.Name -split '\\')[-1] } | 
  Sort-Object -Unique 

$mruOrgIdUserIds = Get-ChildItem "$basePath\*\User MRU\OrgId_*" | 
  foreach { ($_.Name -split '\\')[-1] } | 
  Sort-Object -Unique 

$roamingADUserIds = Get-ChildItem "$basePath\Common\Roaming\Identities\*_AD" | 
  foreach { ($_.Name -split '\\')[-1] } | 
  Sort-Object -Unique 

$roamingOrgIdUserIds = Get-ChildItem "$basePath\Common\Roaming\Identities\*_OrgId" | 
  foreach { ($_.Name -split '\\')[-1] } | 
  Sort-Object -Unique 


## Verify that only one AD and one OrgID settings exist for both MRU and Roaming

if ($mruAdUserIds.Count -ne 1)
{ 
    Write-Host "A single AD identity is required. Not performing migration."
    return 
}

if ($mruOrgIdUserIds.Count -ne 1)
{ 
    Write-Host "A single OrgId identity is required. Not performing migration."
    return 
}

if ($roamingADUserIds.Count -ne 1)
{ 
    Write-Host "A single AD identity is required. Not performing migration."
    return 
}

if ($roamingOrgIdUserIds.Count -ne 1)
{ 
    Write-Host "A single OrgId identity is required. Not performing migration."
    return 
}


## Backup existing OrgId MRU

foreach($item in Get-ChildItem "$basePath\*\User MRU\OrgId_*" -ErrorAction SilentlyContinue)
{
	$adKeyPath = $item.Name
		
	## Find out which app this key is for
	$appName = ($adKeyPath -split '\\')[5]
		 
    $orgIdUserMruPathForReg = $basePathForReg + $appName + "\User MRU\" + $mruOrgIdUserIds

    Write-Host "Backing up MRU from $orgIdUserMruPathForReg"

    reg.exe export $orgIdUserMruPathForReg "$BackupPath\OfficeSettingsMigrateBackup_orgIdUserMru_$appName.reg"

    if ((Test-Path "$BackupPath\OfficeSettingsMigrateBackup_orgIdUserMru_$appName.reg") -eq $false)
    {
        Write-Host "Did not successfully write backup file. Not proceeding with migration."
        return
    }
}


## Backup existing OrgId Roaming

$orgIdUserRoamingPathForReg = $basePathForReg + "Common\Roaming\Identities\" + $roamingOrgIdUserIds

Write-Host "Backing up MRU from $orgIdUserRoamingPathForReg"

reg.exe export $orgIdUserRoamingPathForReg "$BackupPath\OfficeSettingsMigrateBackup_orgIdUserRoaming.reg"

if ((Test-Path "$BackupPath\OfficeSettingsMigrateBackup_orgIdUserRoaming.reg") -eq $false)
{
    Write-Host "Did not successfully write backup file. Not proceeding with migration."
    return
}


## Copy MRU from AD to OrgID

foreach($item in Get-ChildItem "$basePath\*\User MRU\AD_*" -ErrorAction SilentlyContinue)
{
	$adKeyPath = $item.Name
		
	## Find out which app this key is for
	$appName = ($adKeyPath -split '\\')[5]
		
	$adUserMruPath = $basePath + $appName + "\User MRU\" + $mruAdUserIds
    $orgIdUserMruPath = $basePath + $appName + "\User MRU\" + $mruOrgIdUserIds

    Write-Host "Copying MRU from $adUserMruPath to $orgIdUserMruPath"

    Remove-Item $orgIdUserMruPath -Recurse
    Copy-Item $adUserMruPath -Dest $orgIdUserMruPath -Recurse
}


## Copy Roaming from AD to OrgID

$adUserRoamingPath = $basePath + "Common\Roaming\Identities\" + $roamingADUserIds
$orgIdUserRoamingPath = $basePath + "Common\Roaming\Identities\" + $roamingOrgIdUserIds

Write-Host "Copying roaming settings from $adUserRoamingPath to $orgIdUserRoamingPath"

Remove-Item $orgIdUserRoamingPath -Recurse
Copy-Item $adUserRoamingPath -Dest $orgIdUserRoamingPath -Recurse
