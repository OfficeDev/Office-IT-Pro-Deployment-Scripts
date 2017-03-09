[CmdletBinding(SupportsShouldProcess=$true)]
param(
	[Parameter(Mandatory=$True)]
    [string] $UserName
)

Function Reset-OfficeProPlus-ActivationState {
<#
.Synopsis
Script to reset an Office 365 ProPlus 2013/2016  activation/installation to a clean state

.DESCRIPTION
Script to reset an Office 365 ProPlus 2013/2016  activation/installation to a clean state

.PARAMETER UserName
The Name of the user account that needs to be reset
Eg: yourname@domain.com

.EXAMPLE
Reset-OfficeProPlus-ActivationState

.LINK
https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts

.NOTES   
Name: Reset-OfficeProPlus-ActivationState
Version: 1.0.0
DateCreated: 2017-03-06
DateUpdated: 2017-03-08



#>
[CmdletBinding(SupportsShouldProcess=$true)]
param(
	[Parameter(Mandatory=$True)]
    [string] $UserName
)

begin {

	 #ToDo: Replace with a windows temp path later. 
	 $logPath = "D:\Powershell\log.txt"

	 $HKCU = 'HKEY_CURRENT_USER'
	 $HKU = 'HKEY_USERS'
	 
	 $identitiesFolder = 'SOFTWARE\Microsoft\Office\15.0\Common\Identity\Identities',
                         'SOFTWARE\Microsoft\Office\16.0\Common\Identity\Identities'
						 
	 $legacyOSPPvbsFilePath = "'C:\Program Files\Microsoft Office\Office15\ospp.vbs'"
	 
	 $SID = (New-Object System.Security.Principal.NTAccount($UserName)).Translate([System.Security.Principal.SecurityIdentifier]).value
	 
	 $locationstoClear = -join($env:APPDATA,'\Microsoft\Credentials'),
						 -join($env:LOCALAPPDATA ,'\Microsoft\Credentials'),
						 -join($env:APPDATA,'\Microsoft\Protect'),
						 -join($env:LOCALAPPDATA ,'\Microsoft\Office\15.0\Licensing'),
						 -join($env:LOCALAPPDATA ,'\Microsoft\Office\16.0\Licensing')
						 
	 $registrytoClear = 'HKEY_CURRENT_USER\Software\Microsoft\Protected Storage System Provider',
						'HKEY_CURRENT_USER\Software\Microsoft\Office\15.0\Common\Identity',
						'HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\Identity',
						-join('HKEY_USERS\',$SID,'\Software\Microsoft\Office\15.0\Common\Identity'),
						-join('HKEY_USERS\',$SID,'\Software\Microsoft\Office\16.0\Common\Identity')
}


process {	
	#Part 1: Remove Office 365 license for Subscription based installs 
	[string]$Action = "`"cmd /C %systemroot%\system32\cscript.exe $legacyOSPPvbsFilePath /dstatus`""
	
	Invoke-Expression "$Action" -OutVariable out | Tee-Object -FilePath $logPath
	

    #Part 2: Remove cached identities from HKCU registry:
    foreach ($identity in $identitiesFolder) {
	
		#2a. Remove from HKCU
		$fullPath = join-path $HKCU $identity
        
		#Not deleting at the moment to not screw up the local system.
		#To Do: Enable below line and remove the following line
		#Get-ChildItem -Path Registry::$fullPath -Recurse | remove-item -Confirm:$false
        Get-ChildItem -Path Registry::$fullPath
		
		Write-Host "Removed all identities from " $fullPath
		
		#2b. Remove from HKU for Shared Users
		$fullPath = [io.path]::combine($HKU, $SID, $identity)
		
		#Not deleting at the moment to not screw up the local system.
		#To Do: Enable below line and remove the following line
		#Get-ChildItem -Path Registry::$fullPath -Recurse | remove-item -Confirm:$false
        Get-ChildItem -Path Registry::$fullPath
		
		Write-Host "Removed all identities from " $fullPath
		
    }
	
	
	#Part 3: Remove the stored Credentials in the Credential Manager
	$Result = cmdkey /list 
	foreach ($Entry in $Result) 
        { 
            if ($Entry) 
            { 
                $Line = $Entry.Trim(); 
                if ($Line.Contains('Target: ')) 
                { 
					if ($Line.Contains('MicrosoftOffice15') -or $Line.Contains('MicrosoftOffice16')) 
					{
						$Target = $Line.Replace('Target: ',''); 
						
						#Not removing at the moment to not screw up the local system.
						#To Do: Enable below line and remove the following line
						#cmdkey /delete: $Target 
						cmdkey /list: $Target
					}
				}
			}
		}
		
	#Part 4: Persisted locations that must be cleared
	#4a. Registry
	foreach ($folder in $registrytoClear) 
	{
		#Not deleting at the moment to not screw up the local system.
		#To Do: Enable below line and remove the following line
		#Get-ChildItem -Path Registry::$folder -Recurse | remove-item -Confirm:$false
        Get-ChildItem -Path Registry::$folder
		
		Write-Host "Removed all entries from " $folder
	}
	
	#4b. Folders
	foreach ($folder in $locationstoClear) 
	{
		#Not deleting at the moment to not screw up the local system.
		#To Do: Enable below line and remove the following line
		#Get-ChildItem -Path $folder -Recurse | remove-item -Confirm:$false
        Get-ChildItem -Path $folder
		
		Write-Host "Removed all entries from " $folder
	}
}

}

Reset-OfficeProPlus-ActivationState -UserName $UserName