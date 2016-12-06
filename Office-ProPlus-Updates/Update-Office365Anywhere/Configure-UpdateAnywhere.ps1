Function Configure-UpdateAnywhere {
<#
.Synopsis
Configures an existing Group Policy Object (GPO) to schedule a task on workstations to query the update Office using the update anywhere script

.NOTES   
Name: Configure-UpdateAnywhere
Version: 1.0.1

.LINK
https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts

.PARAMETER GpoName
The name of the Group Policy Object (GPO) to configure to inventory Office Clients

.PARAMETER Domain
The Domain name of the target Active Directory Domain

.PARAMETER WaitForUpdateToFinish
If this parameter is set to $true then the function will monitor the Office update and will not exit until the update process has stopped.
If this parameter is set to $false then the script will exit right after the update process has been started.  By default this parameter is set
to $true

.PARAMETER EnableUpdateAnywhere
This parameter controls whether the UpdateAnywhere functionality is used or not. When enabled the update process will check the availbility
of the update source set for the client.  If that update source is not available then it will update the client from the Microsoft Office CDN.
When set to $false the function will only use the Update source configured on the client. By default it is set to $true.

.PARAMETER ForceAppShutdown
This specifies whether the user will be given the option to cancel out of the update. However, if this variable is set to True, then the applications will be shut down immediately and the update will proceed.

.PARAMETER UpdatePromptUser
This specifies whether or not the user will see this dialog before automatically applying the updates:

.PARAMETER DisplayLevel
This specifies whether the user will see a user interface during the update. Setting this to false will hide all update UI (including error UI that is encountered during the update scenario).

.PARAMETER UpdateToVersion
This specifies the version to which Office needs to be updated to.  This can used to install a newer or an older version than what is presently installed.


.EXAMPLE
Configure-UpdateAnywhere -GpoName UpdateGPO

#>
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param
    (
	    [Parameter(Mandatory=$True)]
	    [String]$GpoName,

	    [Parameter()]
	    [String]$Domain = $NULL,

        [Parameter()]
        [bool] $WaitForUpdateToFinish = $true,

        [Parameter()]
        [bool] $EnableUpdateAnywhere = $true,

        [Parameter()]
        [bool] $ForceAppShutdown = $false,

        [Parameter()]
        [bool] $UpdatePromptUser = $false,

        [Parameter()]
        [bool] $DisplayLevel = $false,

        [Parameter()]
        [string] $UpdateToVersion = $NULL
    )

    Begin {
	    $currentExecutionPolicy = Get-ExecutionPolicy
	    Set-ExecutionPolicy Unrestricted -Scope Process -Force  
        $startLocation = Get-Location
    }

    Process {
    try {

    $scriptRoot = GetScriptRoot

    if ($Domain) {
      $Root = [ADSI]"LDAP://$Domain/RootDSE"
    } else {
      $Root = [ADSI]"LDAP://RootDSE"
    }
    
    $DomainPath = $Root.Get("DefaultNamingContext")

    if ($Domain) {
      $gpo = Get-GPO -Name $GpoName -Domain $Domain
    } else {
      $gpo = Get-GPO -Name $GpoName
    }
	
	if(!$gpo -or ($gpo -eq $null))
	{
		Write-Error "The GPO $GpoName could not be found."
        <# write log#>
        $lineNum = Get-CurrentLineNumber    
        $filName = Get-CurrentFileName 
        WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "The GPO $GpoName could not be found."
		Exit
	}

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

    $sourceFileXmlPath = Join-Path $scriptRoot "Files.xml"
    $targetFileXmlPath = Join-Path $filesPath "Files.xml"

    Copy-Item -Path $sourceFileXmlPath -Destination $targetFileXmlPath -Force

    $sourceXmlPath = Join-Path $scriptRoot "ScheduledTasks.xml"
    $targetXmlPath = Join-Path $scriptsPath "ScheduledTasks.xml"

    [System.XML.XMLDocument]$ConfigFile = New-Object System.XML.XMLDocument
    $ConfigFile.Load($sourceXmlPath)
    $argNode = $ConfigFile.SelectSingleNode("/ScheduledTasks/ImmediateTaskV2/Properties/Task/Actions/Exec/Arguments")
    
    $innerText = "-File %Windir%\Temp\Update-Office365Anywhere.ps1 -WaitForUpdateToFinish `$$WaitForUpdateToFinish -EnableUpdateAnywhere `$$EnableUpdateAnywhere -ForceAppShutdown `$$ForceAppShutdown -UpdatePromptUser `$$UpdatePromptUser -DisplayLevel `$$DisplayLevel"
    if([string]::IsNullOrWhiteSpace($UpdateToVersion) -eq $false){
        $innerText = "-File %Windir%\Temp\Update-Office365Anywhere.ps1 -WaitForUpdateToFinish `$$WaitForUpdateToFinish -EnableUpdateAnywhere `$$EnableUpdateAnywhere -ForceAppShutdown `$$ForceAppShutdown -UpdatePromptUser `$$UpdatePromptUser -DisplayLevel `$$DisplayLevel -UpdateToVersion `$$UpdateToVersion"
    } 
    $argNode.InnerText = $innerText
    $ConfigFile.Save($sourceXmlPath)
     
    Copy-Item -Path $sourceXmlPath -Destination $targetXmlPath -Force

    $sourcePsPath = Join-Path $scriptRoot "Update-Office365Anywhere.ps1"
    $targetPsPath = Join-Path $netlogonPath "Update-Office365Anywhere.ps1"
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
            $fileName = $_.InvocationInfo.ScriptName.Substring($_.InvocationInfo.ScriptName.LastIndexOf("\")+1)
            WriteToLogFile -LNumber $_.InvocationInfo.ScriptLineNumber -FName $fileName -ActionError $_
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
    Write-Host "The Group Policy '$GpoName' has been modified to update Office anywhere via Scheduled Task." -BackgroundColor DarkBlue
    Write-Host "Once Group Policy has refreshed as scheduled task will be created to run the scheduled task." -BackgroundColor DarkBlue
    <# write log#>
    $lineNum = Get-CurrentLineNumber    
    $filName = Get-CurrentFileName 
    WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "GPO Modified"
    WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "The Group Policy '$GpoName' has been modified to update Office anywhere via Scheduled Task."
    WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Once Group Policy has refreshed as scheduled task will be created to run the scheduled task."

    } catch {
            $fileName = $_.InvocationInfo.ScriptName.Substring($_.InvocationInfo.ScriptName.LastIndexOf("\")+1)
            WriteToLogFile -LNumber $_.InvocationInfo.ScriptLineNumber -FName $fileName -ActionError $_
    }

    }

    End {
       
       $setLocation = Set-Location $startLocation


    }

}

Function GetScriptRoot() {
 process {
     [string]$scriptPath = "."

     if ($PSScriptRoot) {
       $scriptPath = $PSScriptRoot
     } else {
       $scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
     }

     return $scriptPath
 }
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

function Get-CurrentLineNumber {
    $MyInvocation.ScriptLineNumber
}

function Get-CurrentFileName{
    $MyInvocation.ScriptName.Substring($MyInvocation.ScriptName.LastIndexOf("\")+1)
}

function Get-CurrentFunctionName {
    (Get-Variable MyInvocation -Scope 1).Value.MyCommand.Name;
}