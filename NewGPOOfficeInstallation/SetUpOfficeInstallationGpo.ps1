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
	[String]$ScriptName = "InstallOffice2016.ps1"
)
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

	Set-Location $scriptsPath
	
	#region PSSCripts.ini
	$psScriptsFileName = "psscripts.ini"

	$psScriptsFilePath = ".\$psScriptsFileName"

	$encoding = 'Unicode' #[System.Text.Encoding]::Unicode

	if(!(Test-Path $psScriptsFilePath))
	{
		$baseContent = @()
		$baseContent = $baseContent + " `r`n"
		$baseContent = $baseContent + "[ScriptsConfig]"
		$baseContent = $baseContent + "StartExecutePSFirst=true"
		$baseContent = $baseContent + "[Startup]"
		
		$baseContent | Out-File -FilePath $psScriptsFilePath -Encoding unicode -Force
		
		$file = Get-ChildItem -Path $psScriptsFilePath
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

	$lastScriptIndex = [Convert]::ToInt32($num)

	$nextScriptIndex = $lastScriptIndex + 1

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

