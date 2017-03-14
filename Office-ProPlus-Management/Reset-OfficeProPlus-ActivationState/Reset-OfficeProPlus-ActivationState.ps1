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
DateUpdated: 2017-03-13



#>
[CmdletBinding(SupportsShouldProcess=$true)]
param(
	[Parameter(Mandatory=$True)]
    [string] $UserName
)

begin {

	 #ToDo: Replace with a windows temp path later. 
	 $logPath = -join($env:WINDIR,'\Temp\OSPPCleanupLog.txt')

	 $HKCU = 'HKEY_CURRENT_USER'
	 $HKU = 'HKEY_USERS'

	 
	 $identitiesFolder = 'SOFTWARE\Microsoft\Office\15.0\Common\Identity\Identities',
                         'SOFTWARE\Microsoft\Office\16.0\Common\Identity\Identities'
						 
	 $checkOSPPvbsFilePathOffice15 = -join($env:ProgramW6432,'\Microsoft Office\Office15\ospp.vbs')
     $checkOSPPvbsFilePathx86Office15 = -join(${env:ProgramFiles(x86)},'\Microsoft Office\Office15\ospp.vbs')
     $checkOSPPvbsFilePathOffice16 = -join($env:ProgramW6432,'\Microsoft Office\Office16\ospp.vbs')
     $checkOSPPvbsFilePathx86Office16 = -join(${env:ProgramFiles(x86)},'\Microsoft Office\Office16\ospp.vbs')

     if(Test-Path $checkOSPPvbsFilePathOffice15)
     {
        $legacyOSPPvbsFilePath = $checkOSPPvbsFilePathOffice15
     }
     elseif(Test-Path $checkOSPPvbsFilePathx86Office15)
     {
        $legacyOSPPvbsFilePath = $checkOSPPvbsFilePathx86Office15
     }
     elseif(Test-Path $checkOSPPvbsFilePathOffice16)
     {
        $legacyOSPPvbsFilePath = $checkOSPPvbsFilePathOffice16
     }
     elseif(Test-Path $checkOSPPvbsFilePathx86Office16)
     {
        $legacyOSPPvbsFilePath = $checkOSPPvbsFilePathx86Office16
     }
 
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

    #ToDo: 1. Path validation, 2. Change file path to C:\windows\temp, 3. Parse the output file 
     

	#Part 1: Remove Office 365 license for Subscription based installs 

    #Write results of ospp.vbs to a temp file
    cmd /C cscript $legacyOSPPvbsFilePath /dstatus > $logPath

    $reader = [System.IO.File]::OpenText($logPath)
    try
    {
        while($null -ne ($line = $reader.ReadLine())) {
            if ($line.Contains('Last 5 characters of installed product key: ')) 
            { 
                $target = $line.Replace('Last 5 characters of installed product key: ','');      
            
                #Command to uninstall product key
                cmd /C cscript $legacyOSPPvbsFilePath /unpkey:$target  
            }
        }
    }
    finally
    {
        $reader.Close();
    }


    #Part 2: Remove cached identities from HKCU registry:
    foreach ($identity in $identitiesFolder) {
	
		#2a. Remove from HKCU
		$fullPath = join-path $HKCU $identity
       
		Get-ChildItem -Path Registry::$fullPath -EA SilentlyContinue | remove-item -Confirm:$false -Force -Recurse		
		Write-Host "Removed all identities from " $fullPath
        
		
		#2b. Remove from HKU for Shared Users
		$fullPath = [io.path]::combine($HKU, $SID, $identity)
		
		Get-ChildItem -Path Registry::$fullPath -EA SilentlyContinue | remove-item -Confirm:$false -Force -Recurse
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
						cmdkey /delete: $Target 
						
					}
				}
			}
		}
		
	#Part 4: Persisted locations that must be cleared
	#4a. Registry
	foreach ($folder in $registrytoClear) 
	{
        
		Remove-ItemProperty -Path Registry::$folder -Name * -EA SilentlyContinue
		Write-Host "Removed all entries from " $folder
        
	}
	
	#4b. Folders
	foreach ($folder in $locationstoClear) 
	{
        if(Test-Path $folder)
        {
		    Get-ChildItem -Path $folder -EA SilentlyContinue | remove-item -Confirm:$false -Force -Recurse
		    Write-Host "Removed all entries from " $folder
        }
	}
}

}

Reset-OfficeProPlus-ActivationState -UserName $UserName