Function Configure-GPOOfficeInventory {
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param
    (
	    [Parameter(Mandatory=$True)]
	    [String]$GpoName,

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

    $Root = [ADSI]"LDAP://RootDSE"
    $DomainPath = $Root.Get("DefaultNamingContext")

    Write-Host "Configuring Group Policy to Install Office Click-To-Run"
    Write-Host

    Write-Host "Searching for GPO: $GpoName..." -NoNewline
	$gpo = Get-GPO -Name $GpoName
	
	if(!$gpo -or ($gpo -eq $null))
	{
		Write-Error "The GPO $GpoName could not be found."
		Exit
	}

    Write-Host "GPO Found"
    Write-Host "Modifying GPO: $GpoName..." -NoNewline

	$baseSysVolPath = "$env:LOGONSERVER\sysvol"

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

    }

    End {
       
       $setLocation = Set-Location $startLocation


    }


}

Function Export-GPOOfficeInventory {
    Param (
        [Parameter()]
        [string]$AttributeToStoreOfficeVersion = "info"
    )

    $AttributeToStoreOfficeVersion = $AttributeToStoreOfficeVersion.ToLower()

    $strFilter = "(&(objectCategory=Computer)($AttributeToStoreOfficeVersion=*))"

    $objDomain = New-Object System.DirectoryServices.DirectoryEntry

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