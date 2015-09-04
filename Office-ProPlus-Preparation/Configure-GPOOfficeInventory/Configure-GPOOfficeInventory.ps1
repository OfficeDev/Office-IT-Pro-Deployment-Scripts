Function Configure-GPOOfficeInventory {
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param
    (
	    [Parameter(Mandatory=$True)]
	    [String]$GpoName,

	
	    [Parameter()]
	    [String]$ScriptName = "InstallOffice2013.ps1"
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

    $createDir = [system.io.directory]::CreateDirectory($scriptsPath) 

	$gptIniFileName = "GPT.ini"
	$gptIniFilePath = ".\$gptIniFileName"
   
	Set-Location $scriptsPath

	$encoding = 'Unicode' #[System.Text.Encoding]::Unicode
	
    #$createDir = [system.io.directory]::CreateDirectory($setupExeTargetPath) 	
	#Copy-Item -Path $setupExeSourcePath -Destination $setupExeTargetPath -Force

    $sourceXmlPath = Join-Path $PSScriptRoot "ScheduledTasks.xml"
    $targetXmlPath = Join-Path $scriptsPath "ScheduledTasks.xml"

    Copy-Item -Path $sourceXmlPath -Destination $targetXmlPath

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
        [string]$currentExt = $currentExt.replace("[", "")
        $currentExt = $currentExt.replace("]", "")

        $extSplit = $currentExt.Split('{')

        foreach ($extGuid in $extSplit) {
          if ($extGuid) {
             $addItem = $extList.Add($extGuid.Replace("}", "").ToUpper())
          }
        }
    }

    $extGuids = @("{00000000-0000-0000-0000-000000000000}",`
                  "{CAB54552-DEEA-4691-817E-ED4A4D1AFC72}",`
                  "{AADCED64-746C-4633-A97C-D61349046527}",`
                  "{CAB54552-DEEA-4691-817E-ED4A4D1AFC72}")

    foreach ($extGuid in $extGuids) {
        if (!$extList.Contains($extGuid)) {
          $addItem = $extList.Add($extGuid)
        }
    }

    $newGptExt = "["
    foreach ($extAddGuid in $extList) {
       $newGptExt += "{$extAddGuid}"
    }
    $newGptExt += "]"

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