Function Download-GPOOfficeInstallation {

    [CmdletBinding(SupportsShouldProcess=$true)]
    Param
                        (
	[Parameter(Mandatory=$True)]
	[String]$UncPath,
	
	[Parameter()]
	[String]$Bitness = '32'
    )
    Begin
                {
	$currentExecutionPolicy = Get-ExecutionPolicy
	Set-ExecutionPolicy Unrestricted -Scope Process -Force  
    $startLocation = Get-Location
    }
    Process
    {
	Write-Host 'Updating Config Files'
	
	$setupFileName = 'SetupOffice2013.exe'
	$localSetupFilePath = ".\$setupFileName"
	$setupFilePath = "$UncPath\$localSetupFilePath"
	
	Copy-Item -Path $localSetupFilePath -Destination $UncPath -Force
	
	$downloadConfigFileName = 'Configuration_Download.xml'
	$downloadConfigFilePath = "$UncPath\$downloadConfigFileName"
	$localDownloadConfigFilePath = ".\$downloadConfigFileName"
	
	$installConfigFileName = 'Configuration_InstallLocally.xml'
	$installConfigFilePath = "$UncPath\$installConfigFileName"
	$localInstallConfigFilePath = ".\$installConfigFileName"
	
	$content = [Xml](Get-Content $localDownloadConfigFilePath)
	$addNode = $content.Configuration.Add
	$addNode.OfficeClientEdition = $Bitness
	Write-Host 'Saving Download Configuration XML'	
	$content.Save($downloadConfigFilePath)
	
	$content = [Xml](Get-Content $localInstallConfigFilePath)
	$addNode = $content.Configuration.Add
	$addNode.OfficeClientEdition = $Bitness
	$addNode.SourcePath = $UncPath
	$updatesNode = $content.Configuration.Updates
	$updatesNode.UpdatePath = $UncPath
	Write-Host 'Saving Install Configuration XML'
	$content.Save($installConfigFilePath)
	
	Write-Host 'Setting up Click2Run to download Office to UNC Path'
	
	Set-Location $UncPath
	
	$app = ".\$setupFileName"
	$arguments = "/download", "$downloadConfigFileName"
	
	Write-Host 'Starting Download'
	& $app @arguments
	
	Write-Host 'Download Complete'	
    }
    End
    {
	    Set-ExecutionPolicy $currentExecutionPolicy -Scope Process -Force
        Set-Location $startLocation
    }

}

Function Configure-GPOOfficeInstallation {
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param
    (
	    [Parameter(Mandatory=$True)]
	    [String]$GpoName,
	
	    [Parameter(Mandatory=$True)]
	    [String]$UncPath,
	
	    [Parameter()]
	    [String]$ConfigFileName = "Configuration_InstallLocally.xml",
	
	    [Parameter()]
	    [String]$ScriptName = "InstallOffice2013.ps1"
    )

    Begin
    {
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
	$relativePathToScriptsFolder = "Machine\Scripts"
	$scriptsPath = "{0}\{1}" -f $gpoPath, $relativePathToScriptsFolder

    $createDir = [system.io.directory]::CreateDirectory($scriptsPath) 

	$gptIniFileName = "GPT.ini"
	$gptIniFilePath = ".\$gptIniFileName"
   
	Set-Location $scriptsPath
	
	#region PSSCripts.ini
	$psScriptsFileName = "psscripts.ini"
    $scriptsFileName = "scripts.ini"

	$psScriptsFilePath = ".\$psScriptsFileName"
    $scriptsFilePath = ".\$scriptsFileName"

	$encoding = 'Unicode' #[System.Text.Encoding]::Unicode

	if(!(Test-Path $psScriptsFilePath))
	{
		$baseContent = "`r`n[ScriptsConfig]`r`nStartExecutePSFirst=true`r`n[Startup]"
		$baseContent | Out-File -FilePath $psScriptsFilePath -Encoding unicode -Force
		
		$file = Get-ChildItem -Path $psScriptsFilePath
		$file.Attributes = $file.Attributes -bor ([System.IO.FileAttributes]::Hidden).value__
	}

	if(!(Test-Path $scriptsFilePath))
	{
        "" | Out-File -FilePath $scriptsFilePath -Encoding unicode -Force

		$file = Get-ChildItem -Path $scriptsFilePath
		$file.Attributes = $file.Attributes -bor ([System.IO.FileAttributes]::Hidden).value__
    }
	
	$content = Get-Content -Encoding $encoding -Path $psScriptsFilePath

	$length = $content.Length

	$newContentLength = $length + 2

	$newContent = New-Object System.String[] ($newContentLength)

	$pattern = [string]"\[\w+\]"

	$startUpIndex = 0
	$nextIndex = 0
	$startUpFound = $false

	foreach($s in $content)
	{
		if($s -match $pattern)
		{
		   if($startUpFound)
		   {
			  $nextIndex = $content.IndexOf($s) - 1
			  break
		   }
		   else
		   {
				if($s -eq "[Startup]")
				{
					$startUpIndex = $content.IndexOf($s)
					$startUpFound = $true
				}
		   }
		}
	}

	if($startUpFound -and ($nextIndex -eq 0))
	{
		$nextIndex = $content.Count - 1;
	}
	
	$lastEntry = [string]$content[$nextIndex]

	$num = [regex]::Matches($lastEntry, "\d+")[0].Value   
	
	if($num)
	{
		$lastScriptIndex = [Convert]::ToInt32($num)
	}
	else
	{
		$lastScriptIndex = 0
		$nextScriptIndex = 0
	}
	
	if($lastScriptIndex -gt 0)
	{
		$nextScriptIndex = $lastScriptIndex + 1
	}

	for($i=0; $i -le $nextIndex; $i++)
	{
		$newContent[$i] = $content[$i]
	}

	$newContent[$nextIndex+1] = "{0}CmdLine={1}" -f $nextScriptIndex, $ScriptName

	$newContent[$nextIndex+2] = "{0}Parameters=-UncPath {1} -ConfigFileName {2}" -f $nextScriptIndex, $UncPath, $ConfigFileName

	for($i=$nextIndex; $i -lt $length; $i++)
	{
		$newContent[$i] = $content[$i]
	}

	$newContent | Set-Content -Encoding $encoding -Path $psScriptsFilePath -Force
	#endregion
	
	#region Place the script to attach in the StartUp Folder
	$setupExeSourcePath = "$startLocation\$ScriptName"
	$setupExeTargetPath = "$scriptsPath\StartUp"
    $setupExeTargetPathShutdown = "$scriptsPath\ShutDown"

    $createDir = [system.io.directory]::CreateDirectory($setupExeTargetPath) 
    $createDir = [system.io.directory]::CreateDirectory($setupExeTargetPathShutdown) 
	
	Copy-Item -Path $setupExeSourcePath -Destination $setupExeTargetPath -Force
	#endregion
	
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

    if (!$extList.Contains("42B5FAAE-6536-11D2-AE5A-0000F87571E3")) {
      $addItem = $extList.Add("42B5FAAE-6536-11D2-AE5A-0000F87571E3")
    }

    if (!$extList.Contains("40B6664F-4972-11D1-A7CA-0000F87571E3")) {
      $addItem = $extList.Add("40B6664F-4972-11D1-A7CA-0000F87571E3")
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
    Write-Host "The Group Policy '$GpoName' has been modified to install Office at Workstation Startup." -BackgroundColor DarkBlue
    Write-Host "Once Group Policy has refreshed on the Workstations then Office will install on next startup if the computer has access to the Network Share." -BackgroundColor DarkBlue

    }

    End {
       
       $setLocation = Set-Location $startLocation


    }


}