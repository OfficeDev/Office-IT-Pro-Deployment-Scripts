[CmdletBinding(SupportsShouldProcess=$true)]
param(
	[Parameter()]
    [string] $UserName,

    [Parameter()]
    [string] $LogPath = "$env:windir\Temp\OfficeActivationCleanup.log"
)

Function Reset-OfficeProPlusActivationState {
<#
.Synopsis
Script to reset an Office 365 ProPlus 2013/2016 activation/installation to a clean state

.DESCRIPTION
Script to reset an Office 365 ProPlus 2013/2016 activation/installation to a clean state

.PARAMETER UserName
The Name of the user account that needs to be reset
Eg: yourname@domain.com

.PARAMETER LogPath
The full path to the log file. The default path is %windir%\Temp\OfficeActivationCleanup.log

.EXAMPLE
Reset-OfficeProPlusActivationState -UserName karenb@contoso.com
The Office 365 ProPlus activation key will be removed for karenb@contoso.com using the default log path.

.EXAMPLE
Reset-OfficeProPlusActivationState -UserName karenb@contoso.com -LogPath "$env:temp\O365Cleanup.log"
The Office 365 ProPlus activation key will be removed for karenb@contoso.com and logged to "$env:temp\O365Cleanup.log"

.LINK
https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts

.NOTES   
Name: Reset-OfficeProPlusActivationState
Version: 1.0.0
Created: 2017-03-06
Updated: 2017-03-15
#>
[CmdletBinding(SupportsShouldProcess=$true)]
param(
	[Parameter(Mandatory=$True)]
    [string] $UserName,

    [Parameter()]
    [string] $LogPath = "$env:windir\Temp\OfficeActivationCleanup.log"
)

begin {
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

    $HCUregistrytoClear = 'Software\Microsoft\Protected Storage System Provider',
						  'Software\Microsoft\Office\15.0\Common\Identity',
						  'Software\Microsoft\Office\16.0\Common\Identity'

    $HUregistrytoClear = 'Software\Microsoft\Office\15.0\Common\Identity',
                         'Software\Microsoft\Office\16.0\Common\Identity'
    
}

process {	    
    # Part 1: Remove Office 365 license for Subscription based installs 

    # Write results of ospp.vbs to a temp file
    $OSPPContentPath = "$env:TEMP\OSPPContent.txt"
    cmd /C cscript $legacyOSPPvbsFilePath /dstatus > $OSPPContentPath

    $ProductID = "PRODUCT ID"
    $SkuId = "SKU ID"
    $LicenseName = "LICENSE NAME"
    $LicenseDescription = "LICENSE DESCRIPTION"
    $LicenseStatus = "LICENSE STATUS"
    $Last5Chars = "Last 5 characters of installed product key"

    $ProductIDContent = Get-Content $OSPPContentPath | Select-String -Pattern $ProductID
    $SkuIdContent = Get-Content $OSPPContentPath | Select-String -Pattern $SkuId   
    $LicenseNameContent = Get-Content $OSPPContentPath | Select-String -Pattern $LicenseName
    $LicenseDescriptionContent = Get-Content $OSPPContentPath | Select-String -Pattern $LicenseDescription
    $LicenseStatusContent = Get-Content $OSPPContentPath | Select-String -Pattern $LicenseStatus
    $Last5CharsContent = Get-Content $OSPPContentPath | Select-String -Pattern $Last5Chars

    if($ProductIDContent -ne $null){
        writelog $ProductIDContent
        writelog $SkuIdContent
        writelog $LicenseNameContent
        writelog $LicenseDescriptionContent
        writelog $LicenseStatusContent
        writelog $Last5CharsContent

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

        Write-Host ""
        Write-Host "Removing remaining cached files..."

        #Part 2: Remove cached identities from HKCU registry:
        foreach ($identity in $identitiesFolder) {
	    
	    	#2a. Remove from HKCU
	    	$fullPath = Join-Path $HKCU $identity
           
	    	$HKCUIdentities = Get-ChildItem -Path Registry::$fullPath -EA SilentlyContinue | Remove-Item -Confirm:$false -Force -Recurse		
            if($HKCUIdentities -ne $null){		
                Write-Host "`tRemoved all identities from " $fullPath
                writelog "Removed all identities from $fullPath"
            }        
	    	
	    	#2b. Remove from HKU for Shared Users
	    	$fullPath = [io.path]::combine($HKU, $SID, $identity)
	    	
	    	$HKUIdentities = Get-ChildItem -Path Registry::$fullPath -EA SilentlyContinue | Remove-Item -Confirm:$false -Force -Recurse
            if($HKUIdentities -ne $null){		
                Write-Host "`tRemoved all identities from " $fullPath
                writelog "Removed all identities from $fullPath"
            }	
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
	    foreach ($folder in $HCUregistrytoClear) 
	    {
            $folderPath = Join-Path $HKCU $folder
            if(Test-Path Registry::$folderPath)
            {
	    	    Remove-ItemProperty -Path Registry::$folderPath -Name * -EA SilentlyContinue
	    	    Write-Host "`tRemoved all entries from " $folderPath
                writelog "Removed all entries from $folderPath"
            }
            
	    }

        foreach ($folder in $HUregistrytoClear)
        {
            $SIDPath = Join-Path $SID $folder
            $folderPath = Join-Path $HKU $SIDPath
            if(Test-Path Registry::$folderPath)
            {
	    	    Remove-ItemProperty -Path Registry::$folderPath -Name * -EA SilentlyContinue
	    	    Write-Host "`tRemoved all entries from " $folderPath
                writelog "Removed all entries from $folderPath"
            }
        }
	    
	    #4b. Folders
	    foreach ($folder in $locationstoClear) 
	    {
            if(Test-Path $folder)
            {
	    	    Get-ChildItem -Path $folder -EA SilentlyContinue | remove-item -Confirm:$false -Force -Recurse
	    	    Write-Host "`tRemoved all entries from " $folder
                writelog "Removed all entries from $folder"
            }
	    }
    }
    else
    {
        Write-Host "Activation ID for $UserName not found"
        writelog "Activation ID for $UserName not found"
    }
}

}

Function IsDotSourced() {
  [CmdletBinding(SupportsShouldProcess=$true)]
  param(
    [Parameter(ValueFromPipelineByPropertyName=$true)]
    [string]$InvocationLine = ""
  )
  $cmdLine = $InvocationLine.Trim()
  Do {
    $cmdLine = $cmdLine.Replace(" ", "")
  } while($cmdLine.Contains(" "))

  $dotSourced = $false
  if ($cmdLine -match '^\.\\') {
     $dotSourced = $false
  } else {
     $dotSourced = ($cmdLine -match '^\.')
  }

  return $dotSourced
}

function writelog([string]$value = ""){
    $LogOutput = ("$value")
    Out-File -InputObject $LogOutput -FilePath $LogPath -Append -Encoding UTF8
}

$dotSourced = IsDotSourced -InvocationLine $MyInvocation.Line

if (!($dotSourced)) {
    Reset-OfficeProPlusActivationState -UserName $UserName -LogPath $LogPath
}