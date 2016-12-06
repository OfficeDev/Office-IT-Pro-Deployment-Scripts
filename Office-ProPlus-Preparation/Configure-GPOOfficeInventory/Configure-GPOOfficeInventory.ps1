Function Configure-GPOOfficeInventory {
<#
.Synopsis
Configures an existing Group Policy Object (GPO) to schedule a task on workstations to query the version of Office that is installed
on the computer and write that information to an attribute on the computer object in Active Directory.

.DESCRIPTION
If you don't have System Center Configruration Manager (SCCM) or an equivalent software management system then using this script
will provide the capability to inventory what versions of Office are installed in the domain.

.NOTES   
Name: Configure-GPOOfficeInventory 
Version: 1.0.1
DateCreated: 2015-08-20
DateUpdated: 2015-09-04

.LINK
https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts

.PARAMETER GpoName
The name of the Group Policy Object (GPO) to configure to inventory Office Clients

.PARAMETER Domain
The Domain name of the target Active Directory Domain

.PARAMETER AttributeToStoreOfficeVersion
The parameter on the computer object that will store the version of Active Directory.  In order for this script to work 
the computer object's SELF must have write permissions to the attribute specified.  By default a computer in Active Directory
has permissions to write to attributes that are classified as 'Personal Information'.  This functionality is what allows this 
Inventory functionality to work.  The scheduled task that runs on the computer runs under the 'System' context which gives it
permissions to write to its own computer account in Active Directory.  If you would like to use an attribute that is not in the 
'Personal Information' list then you would have to give 'Self' permissions to write to that Attribute on computer object in 
Active Directory.  A list of possible attributes that you can use are listed below.  The default attribute that is used by
this script is Info.  It is an attribute that is unlikely to be already used.  The drawback to using it is that you can 
not see the value in the computer list view in Active Directory Users and computers.

    -info
    -physicalDeliveryOfficeName
    -assistant
    -facsimileTelephoneNumber
    -InternationalISDNNumber
    -personalTitle
    -otherIpPhone
    -ipPhone
    -primaryInternationalISDNNumber
    -thumbnailPhoto
    -postalCode
    -preferredDeliveryMethod
    -registeredAddress
    -streetAddress
    -telephoneNumber
    -teletexTerminalIdentifier
    -telexNumber
    -primaryTelexNumber

.PARAMETER OverWriteFile
Will parameter controls whether or not the Office inventory script will overwrite the Active Directory computer attribute if 
a value already exists for that attribute

.EXAMPLE
Configure-GPOOfficeInventory -GpoName OfficeInventoryGPO

Description:
This Example will configure the GPO 'OfficeInventoryGPO' to inventory the Office version of the workstations to which the Group 
Policy is applied

#>
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param
    (
	    [Parameter(Mandatory=$True)]
	    [String]$GpoName,

	    [Parameter()]
	    [String]$Domain = $NULL,

        [Parameter()]
        [string]$AttributeToStoreOfficeVersion = "info",

        [Parameter()]
        [bool]$OverWriteFile = $false
    )

    Begin {
	    $currentExecutionPolicy = Get-ExecutionPolicy
	    Set-ExecutionPolicy Unrestricted -Scope Process -Force  
        $startLocation = Get-Location
    }

    Process {

    if ($Domain) {
      $Root = [ADSI]"LDAP://$Domain/RootDSE"
    } else {
      $Root = [ADSI]"LDAP://RootDSE"
    }
    
    $DomainPath = $Root.Get("DefaultNamingContext")

    Write-Host "Configuring Group Policy to Inventory Office Clients"
    Write-Host
    <# write log#>
    $lineNum = Get-CurrentLineNumber    
    $filName = Get-CurrentFileName 
    WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Configuring Group Policy to Inventory Office Clients"

    Write-Host "Searching for GPO: $GpoName..." -NoNewline
    <# write log#>
    $lineNum = Get-CurrentLineNumber    
    $filName = Get-CurrentFileName 
    WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Searching for GPO: $GpoName..."
    if ($Domain) {
      $gpo = Get-GPO -Name $GpoName -Domain $Domain
    } else {
      $gpo = Get-GPO -Name $GpoName
    }
	
	if(!$gpo -or ($gpo -eq $null))
	{
    <# write log#>
        $lineNum = Get-CurrentLineNumber    
        $filName = Get-CurrentFileName 
        WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "The GPO $GpoName could not be found."
		Write-Error "The GPO $GpoName could not be found."
		Exit
	}

    Write-Host "GPO Found"
    Write-Host "Modifying GPO: $GpoName..." -NoNewline
    <# write log#>
    $lineNum = Get-CurrentLineNumber    
    $filName = Get-CurrentFileName 
    WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "GPO Found, Modifying GPO: $GpoName..."

	$baseSysVolPath = "\\$Domain\sysvol"

	$domain = $gpo.DomainName
    $gpoId = $gpo.Id.ToString()

    $adGPO = [ADSI]"LDAP://CN={$gpoId},CN=Policies,CN=System,$DomainPath"
    	
	$gpoPath = "{0}\{1}\Policies\{{{2}}}" -f $baseSysVolPath, $domain, $gpoId
	$relativePathToSchedTaskFolder = "Machine\Preferences\ScheduledTasks"
	$scriptsPath = "{0}\{1}" -f $gpoPath, $relativePathToSchedTaskFolder
    [system.io.directory]::CreateDirectory($scriptsPath) | Out-Null
    
    $relativePathToFileFolder = "Machine\Preferences\Files"
	$filesPath = "{0}\{1}" -f $gpoPath, $relativePathToFileFolder
    [system.io.directory]::CreateDirectory($filesPath) | Out-Null

    $netlogonPath = "{0}\{1}\Scripts" -f $baseSysVolPath, $domain

	$gptIniFileName = "GPT.ini"
	$gptIniFilePath = ".\$gptIniFileName"
   
	Set-Location $scriptsPath

    $sourceFileXmlPath = Join-Path $PSScriptRoot "Files.xml"
    $targetFileXmlPath = Join-Path $filesPath "Files.xml"

    Copy-Item -Path $sourceFileXmlPath -Destination $targetFileXmlPath -Force

    $sourceXmlPath = Join-Path $PSScriptRoot "ScheduledTasks.xml"
    $targetXmlPath = Join-Path $scriptsPath "ScheduledTasks.xml"

    [string]$overWriteTextValue = "false"
    if ($OverWriteFile) {
       $overWriteTextValue = "true"
    }

    [System.XML.XMLDocument]$ConfigFile = New-Object System.XML.XMLDocument
    $ConfigFile.Load($sourceXmlPath)
    $argNode = $ConfigFile.SelectSingleNode("/ScheduledTasks/ImmediateTaskV2/Properties/Task/Actions/Exec/Arguments")
    $argNode.InnerText = "-File %Windir%\Temp\Inventory-OfficeVersion.ps1 -AttributeToStoreOfficeVersion $AttributeToStoreOfficeVersion -OverWriteAttributeValue $overWriteTextValue"
    $ConfigFile.Save($sourceXmlPath)
     
    Copy-Item -Path $sourceXmlPath -Destination $targetXmlPath -Force

    $sourcePsPath = Join-Path $PSScriptRoot "Inventory-OfficeVersion.ps1"
    $targetPsPath = Join-Path $netlogonPath "Inventory-OfficeVersion.ps1"
    Copy-Item -Path $sourcePsPath -Destination $targetPsPath -Force

	#region Update GPT.ini
	Set-Location $gpoPath   

	$encoding = 'ASCII' #[System.Text.Encoding]::ASCII
	$gptIniContent = Get-Content -Encoding $encoding -Path $gptIniFilePath
	
    [int]$newVersion = 0
	foreach($s in $gptIniContent)
	{
		if($s.StartsWith("Version"))
		{
			$index = $gptIniContent.IndexOf($s)
			#Write-Host "Old GPT.ini Version: $s"

			$num = ($s -split "=")[1]
			$ver = [Convert]::ToInt32($num)
			$newVer = $ver + 1
			$s = $s -replace $num, $newVer.ToString()

			#Write-Host "New GPT.ini Version: $s"
            $newVersion = $s.Split('=')[1]
			$gptIniContent[$index] = $s
			break
		}
	}

    [System.Collections.ArrayList]$extList = New-Object System.Collections.ArrayList

    Try {
       $currentExt = $adGPO.get('gPCMachineExtensionNames')
    } Catch { 

    }

    if ($currentExt) {
        $extSplit = $currentExt.Split(']')

        foreach ($extGuid in $extSplit) {
          if ($extGuid) {
            if ($extGuid.Length -gt 0) {
                $addItem = $extList.Add($extGuid.Replace("[", "").ToUpper())
            }
          }
        }
    }

    $extGuids = @("{00000000-0000-0000-0000-000000000000}{3BAE7E51-E3F4-41D0-853D-9BB9FD47605F}{CAB54552-DEEA-4691-817E-ED4A4D1AFC72}",`
                  "{7150F9BF-48AD-4DA4-A49C-29EF4A8369BA}{3BAE7E51-E3F4-41D0-853D-9BB9FD47605F}",`
                  "{AADCED64-746C-4633-A97C-D61349046527}{CAB54552-DEEA-4691-817E-ED4A4D1AFC72}")


    foreach ($extGuid in $extGuids) {
        if (!$extList.Contains($extGuid)) {
          $addItem = $extList.Add($extGuid)
        }
    }

    foreach ($extAddGuid in $extList) {
       $newGptExt += "[$extAddGuid]"
    }

    $adGPO.put('versionNumber',$newVersion)
    $adGPO.put('gPCMachineExtensionNames',$newGptExt)
    $adGPO.CommitChanges()

    
	$gptIniContent | Set-Content -Encoding $encoding -Path $gptIniFilePath -Force
	
    Write-Host "GPO Modified"
    Write-Host ""
    Write-Host "The Group Policy '$GpoName' has been modified to inventory Office via Scheduled Task." -BackgroundColor DarkBlue
    Write-Host "Once Group Policy has refreshed as scheduled task will be created to run the scheduled task." -BackgroundColor DarkBlue

    <# write log#>
    $lineNum = Get-CurrentLineNumber    
    $filName = Get-CurrentFileName 
    WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "GPO Modified"
    WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "The Group Policy '$GpoName' has been modified to inventory Office via Scheduled Task."
    WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Once Group Policy has refreshed as scheduled task will be created to run the scheduled task."

    }

    End {
       
       $setLocation = Set-Location $startLocation


    }

}

Function Export-GPOOfficeInventory {
<#
.Synopsis
Exports a list of computer and the inventoried Office Versions that were collected by running the function Configure-GPOOfficeInventory
from Active Directory

.DESCRIPTION
After using the function Configure-GPOOfficeInventory to collect the Office Versions into Active Directory this function is used to
export that data.

.NOTES   
Name: Export-GPOOfficeInventory
Version: 1.0.1
DateCreated: 2015-09-03
DateUpdated: 2015-09-04

.LINK
https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts

.PARAMETER Domain
The Domain name of the target Active Directory Domain

.PARAMETER AttributeToStoreOfficeVersion
This attribute has the be the attribute that was used with the Configure-GPOOfficeInventory function.  If no attribute was 
specified then the default attribute 'info' will be used

    -info
    -physicalDeliveryOfficeName
    -assistant
    -facsimileTelephoneNumber
    -InternationalISDNNumber
    -personalTitle
    -otherIpPhone
    -ipPhone
    -primaryInternationalISDNNumber
    -thumbnailPhoto
    -postalCode
    -preferredDeliveryMethod
    -registeredAddress
    -streetAddress
    -telephoneNumber
    -teletexTerminalIdentifier
    -telexNumber
    -primaryTelexNumber

.EXAMPLE
Export-GPOOfficeInventory

Description:
This example will list all of the computers in the domain that have had their Office Versions inventoried by the process
created by the Configure-GPOOfficeInventory function

#>
    Param (
        [Parameter()]
        [string]$Domain = $NULL,

        [Parameter()]
        [string]$AttributeToStoreOfficeVersion = "info"
    )

    $AttributeToStoreOfficeVersion = $AttributeToStoreOfficeVersion.ToLower()

    $strFilter = "(&(objectCategory=Computer)($AttributeToStoreOfficeVersion=*))"

    if ($Domain) {
      $Root = [ADSI]"LDAP://$Domain/RootDSE"
      $DomainPath = $Root.Get("DefaultNamingContext")
      $rootPath = "LDAP://" + $Domain + "/" + $DomainPath
      $objDomain = New-Object System.DirectoryServices.DirectoryEntry($rootPath)
    } else {
      $objDomain = New-Object System.DirectoryServices.DirectoryEntry
    }  

    $objSearcher = New-Object System.DirectoryServices.DirectorySearcher
    $objSearcher.SearchRoot = $objDomain
    $objSearcher.PageSize = 1000
    $objSearcher.Filter = $strFilter
    $objSearcher.SearchScope = "Subtree"

    $colProplist = @("name", "operatingSystem", "distinguishedname", $AttributeToStoreOfficeVersion)
    foreach ($i in $colPropList){
        $objSearcher.PropertiesToLoad.Add($i) | out-Null
    }

    $colResults = $objSearcher.FindAll()

    $results = new-object PSObject[] 0;
    foreach ($objResult in $colResults) {
        $objItem = $objResult.Properties;

        #$objItem.distinguishedname

        $cltr = New-Object -TypeName PSObject
        $cltr | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $objItem.name[0]
        $cltr | Add-Member -MemberType NoteProperty -Name "OperatingSystem" -Value $objItem.operatingsystem[0]
        $cltr | Add-Member -MemberType NoteProperty -Name "OfficeVersion" -Value $objItem.$AttributeToStoreOfficeVersion[0]

        $results += $cltr
    }
    $results
}


function Get-CurrentLineNumber {
    $MyInvocation.ScriptLineNumber
}

function Get-CurrentFileName{
    $MyInvocation.ScriptName.Substring($MyInvocation.ScriptName.LastIndexOf("\")+1)
}

function Get-CurrentFunctionName {
    (Get-Variable MyInvocation -Scope 1).Value.MyCommand.Name;
}

Function WriteToLogFile() {
    param( 
      [Parameter(Mandatory=$true)]
      [string]$LNumber,
      [Parameter(Mandatory=$true)]
      [string]$FName,
      [Parameter(Mandatory=$true)]
      [string]$ActionError
    )
    try{
        $headerString = "Time".PadRight(30, ' ') + "Line Number".PadRight(15,' ') + "FileName".PadRight(60,' ') + "Action"
        $stringToWrite = $(Get-Date -Format G).PadRight(30, ' ') + $($LNumber).PadRight(15, ' ') + $($FName).PadRight(60,' ') + $ActionError

        #check if file exists, create if it doesn't
        $getCurrentDatePath = "C:\Windows\Temp\" + (Get-Date -Format u).Substring(0,10)+"OfficeAutoScriptLog.txt"
        if(Test-Path $getCurrentDatePath){#if exists, append    
             Add-Content $getCurrentDatePath $stringToWrite
        }
        else{#if not exists, create new
             Add-Content $getCurrentDatePath $headerString
             Add-Content $getCurrentDatePath $stringToWrite
        }
    } catch [Exception]{
        Write-Host $_
    }
}