[CmdletBinding(SupportsShouldProcess=$true)]
Param
(
	[Parameter(Mandatory=$true)]
	[String]$GpoName,

    [Parameter(Mandatory=$true)]
    [string]$UncPath,

    [Parameter(Mandatory=$true)]
    [string]$CommonFileShare,
	
	[Parameter()]
	[String]$ScriptName = "Deploy-TelemetryAgent.ps1"

)

<#

.SYNOPSIS
Adds the Deploy-TelemetryAgent and parameters to the GPO specified.

.DESCRIPTION
Given a GpoName and UncPath the Deploy-TelemetryAgent.ps1 script will
be added to the GPO with parameters pointing to the shared folder
containing the telemetry agent msi files.

.PARAMETER GpoName
The name of the GPO to be customized.

.PARAMETER UncPath
The path of the shared drive hosting the osmia32 and osmia64 msi 
files. These are the installation files for the telemetry agent.

.EXAMPLE
./Set-TelemetryStartup -GpoName "Office Telemetry" -UncPath "\\Server1\Sharedfolder"
The Deploy-TelemetryAgent.ps1 script will be added to the Startup folder
in a GPO named Office Telemetry and will create parameters for the UNC path.

#>

Begin
{
	$currentExecutionPolicy = Get-ExecutionPolicy
	Set-ExecutionPolicy Unrestricted -Scope Process -Force  
    $startLocation = Get-Location
}
Process
{
	$gpo = Get-GPO -Name $GpoName
	
	if(!$gpo -or ($gpo -eq $null))
	{
		Write-Error "The GPO $GpoName could not be found."
		Exit
	}

	$baseSysVolPath = "$env:LOGONSERVER\sysvol"

	$domain = $gpo.DomainName

	$gpoId = $gpo.Id.ToString()
	$gpoPath = "{0}\{1}\Policies\{{{2}}}" -f $baseSysVolPath, $domain, $gpoId
	$relativePathToScriptsFolder = "Machine\Scripts"
	$scriptsPath = "{0}\{1}" -f $gpoPath, $relativePathToScriptsFolder

	$gptIniFileName = "GPT.ini"
	$gptIniFilePath = ".\$gptIniFileName"

	
    if(!(Test-Path -Path $scriptsPath)){
        New-Item -ItemType Directory -Path "$gpoPath\Machine\Scripts\Startup" -Force
        }

    Set-Location $scriptsPath
	
	#region PSSCripts.ini
	$psScriptsFileName = "psscripts.ini"

	$psScriptsFilePath = ".\$psScriptsFileName"

	$encoding = 'Unicode' #[System.Text.Encoding]::Unicode

	if(!(Test-Path $psScriptsFilePath))
	{
		
		$baseContent = "`r`n[ScriptsConfig]`r`nStartExecutePSFirst=true`r`n[Startup]"
		
		$baseContent | Out-File -FilePath $psScriptsFilePath -Encoding unicode -Force
		
		$file = Get-ChildItem -Path $psScriptsFilePath -Force
		$file.Attributes = $file.Attributes -bor ([System.IO.FileAttributes]::Hidden).value__
	}
	
	$content = Get-Content -Encoding $encoding -Path $psScriptsFilePath -Force

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

    $newContent[$nextIndex+2] = "{0}Parameters=-UncPath {1} -CommonFileShare {2}" -f $nextScriptIndex, $UncPath, $CommonFileShare
    
	
	for($i=$nextIndex; $i -lt $length; $i++)
	{
		$newContent[$i] = $content[$i]
	}

	$newContent | Set-Content -Encoding $encoding -Path $psScriptsFilePath -Force
	#endregion
	
	#region Place the script to attach in the StartUp Folder
	$setupExeSourcePath = "$startLocation\$ScriptName"
	$setupExeTargetPath = "$scriptsPath\StartUp"
	
	Copy-Item -Path $setupExeSourcePath -Destination $setupExeTargetPath -Force
	#endregion
	
	#region Update GPT.ini
	Set-Location $gpoPath   

	$encoding = 'UTF8' #[System.Text.Encoding]::UTF
	$gptIniContent = Get-Content -Encoding $encoding -Path $gptIniFilePath
	
	foreach($s in $gptIniContent)
	{
		if($s.StartsWith("Version"))
		{
			$index = $gptIniContent.IndexOf($s)

			Write-Host "Old GPT.ini Version: $s"

			$num = ($s -split "=")[1]

			$ver = [Convert]::ToInt32($num)

			$newVer = $ver + 1

			$s = $s -replace $num, $newVer.ToString()

			Write-Host "New GPT.ini Version: $s"

			$gptIniContent[$index] = $s
			break
		}
	}

	$gptIniContent | Set-Content -Encoding $encoding -Path $gptIniFilePath -Force
	#endregion		
}
End
{
	Set-ExecutionPolicy $currentExecutionPolicy -Scope Process -Force
    Set-Location $startLocation
}

