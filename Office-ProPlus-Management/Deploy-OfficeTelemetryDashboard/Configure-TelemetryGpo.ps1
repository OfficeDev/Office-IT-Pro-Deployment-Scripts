Param
(
    [Parameter(Mandatory=$true)]
    [string]$GpoName,

    [Parameter(Mandatory=$true)]
    [string]$TelemetryServer,

    [Parameter()]
    [string]$CommonFileShare = $NULL,

    [Parameter()]
    [string]$AgentShare = $NULL,

    [Parameter()]
    [string]$Domain = $NULL,
    
    [Parameter()]
    [string]$OfficeVersion,

	[Parameter()]
	[String]$ScriptName = "Deploy-TelemetryAgent.ps1"
)

function Set-TelemetryStartup() {
        [CmdletBinding(SupportsShouldProcess=$true)]
        Param
        (
	        [Parameter(Mandatory=$true)]
	        [String]$GpoName,

            [Parameter(Mandatory=$true)]
            [string]$TelemetryServer,

            [Parameter()]
            [string]$CommonFileShare = $NULL,

            [Parameter()]
            [string]$AgentShare = $NULL,
	
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
        ./Set-TelemetryStartup -GpoName "Office Telemetry" -CommonFileShare "\\Server1\TDShared"
        The Deploy-TelemetryAgent.ps1 script will be added to the Startup folder
        in a GPO named Office Telemetry and will create parameters for the UNC path.

        .EXAMPLE
        ./Set-TelemetryStartup -GpoName "Office Telemetry" -CommonFileShare "\\Server1\TDShared" -agentShare "\\Server2\Telemetry Agent"
        The Deploy-TelemetryAgent.ps1 script will be added to the Startup folder
        in a GPO named Office Telemetry and will create parameters for the file share the 
        agent will upload data to and the agent share containing the agent msi files.

        #>

        Begin
        {
	        $currentExecutionPolicy = Get-ExecutionPolicy
	        Set-ExecutionPolicy Unrestricted -Scope Process -Force  
            $startLocation = Get-Location

            if ($TelemetryServer) {
               if (!($CommonFileShare)) {
                  $CommonFileShare = "\\$TelemetryServer\TDShared"
               }
               if (!($AgentShare)) {
                  $AgentShare = "\\$TelemetryServer\TelemetryAgent"
               }
            }
        }
        Process
        {

            Write-Host

            if (!(Test-Path -Path $CommonFileShare)) {
              throw "Common File Share Path Not Available: $CommonFileShare"
            }

            if (!(Test-Path -Path $AgentShare)) {
              throw "Common File Share Path Not Available: $AgentShare"
            }

            $Root = [ADSI]"LDAP://RootDSE"
            $DomainPath = $Root.Get("DefaultNamingContext")

            Write-Host "Configuring Group Policy to Install Telemetry Agent"
            <# write log#>
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Configuring Group Policy to Install Telemetry Agent"

	        $gpo = Get-GPO -Name $GpoName
	
	        if(!$gpo -or ($gpo -eq $null))
	        {
                <# write log#>
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "The GPO $GpoName could not be found."
		        Write-Error "The GPO $GpoName could not be found."
		        Exit
	        }

	        $baseSysVolPath = "$env:LOGONSERVER\sysvol"

	        $domain = $gpo.DomainName
	        $gpoId = $gpo.Id.ToString()

            $adGPO = [ADSI]"LDAP://CN={$gpoId},CN=Policies,CN=System,$DomainPath"

	        $gpoPath = "{0}\{1}\Policies\{{{2}}}" -f $baseSysVolPath, $domain, $gpoId
	        $relativePathToScriptsFolder = "Machine\Scripts"
	        $scriptsPath = "{0}\{1}" -f $gpoPath, $relativePathToScriptsFolder

	        $gptIniFileName = "GPT.ini"
	        $gptIniFilePath = ".\$gptIniFileName"

	
            if(!(Test-Path -Path $scriptsPath)){
                New-Item -ItemType Directory -Path "$gpoPath\Machine\Scripts\Startup" -Force | Out-Null
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

            $newContent[$nextIndex+2] = "{0}Parameters=-CommonFileShare {1} -agentShare {2}" -f $nextScriptIndex, $CommonFileShare, $agentShare
    
	
	        for($i=$nextIndex; $i -lt $length; $i++)
	        {
		        $newContent[$i] = $content[$i]
	        }

	        $newContent | Set-Content -Encoding $encoding -Path $psScriptsFilePath -Force
	        #endregion
	
	        #region Place the script to attach in the StartUp Folder
	        $setupExeSourcePath = "$startLocation\$ScriptName"
	        $setupExeTargetPath = "$scriptsPath\StartUp"
	
	        Copy-Item -Path $setupExeSourcePath -Destination $setupExeTargetPath -Force | Out-Null
	        #endregion
	
	        #region Update GPT.ini
	        Set-Location $gpoPath   

	        $encoding = 'UTF8' #[System.Text.Encoding]::UTF
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
	        #endregion		
        }
        End
        {
            Set-Location $startLocation
        }


   }

<#

.SYNOPSIS
Create the Telemetry GPO on the Domain Controller

.DESCRIPTION
Creates a group policy that that specifies the 
Telemetry agent file share location and allows
the agent to log and upload.

.PARAMETER GpoName
The name of the GPO to be created.

.PARAMETER CommonFileShare
The name of the Shared Drive hosting the telemetry database.

.PARAMETER OfficeVersion
The version of office used in your environment. If a version
earlier than 2013 is used do not use this parameter.

.EXAMPLE
./Create-TelemetryGpo -GpoName "Office Telemetry" -CommonFileShare "Server1" -officeVersion 2013
A GPO named "Office Telemetry" will be created. Registry keys will be
created to enable telemetry agent logging, uploading, and the commonfileshare 
path set to \\Server1\TDShared. 

.EXAMPLE
./Create-TelemetryGpo -GpoName "Office Telemetry"
A GPO named "Office Telemetry" will be created.

#>
    Write-Host
    
    Import-Module -Name grouppolicy

    if ($Domain) {
      $existingGPO = Get-GPO -Name $gpoName -Domain $Domain -ErrorAction SilentlyContinue
    } else {
      $existingGPO = Get-GPO -Name $gpoName -ErrorAction SilentlyContinue
    }

    if ($TelemetryServer) 
    {
        if (!($CommonFileShare)) 
        {
            $CommonFileShare = "\\$TelemetryServer\TDShared"
        }
    }
 
    if (!($existingGPO)) 
    {
        Write-Host "Creating a new Group Policy..."
        <# write log#>
        $lineNum = Get-CurrentLineNumber    
        $filName = Get-CurrentFileName 
        WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Creating a new Group Policy..."

        if ($Domain) {
          New-GPO -Name $gpoName -Domain $Domain
        } else {
          New-GPO -Name $gpoName
        }
    } else {
       Write-Host "Group Policy Already Exists..."
        <# write log#>
        $lineNum = Get-CurrentLineNumber    
        $filName = Get-CurrentFileName 
        WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Group Policy Already Exists..."
    }

    #The same share created in Deploy-TelemetryDashboard.ps1
    $shareName = "TDShared"
    
    Write-Host "Configuring Group Policy '$gpoName': " -NoNewline
<# write log#>
    $lineNum = Get-CurrentLineNumber    
    $filName = Get-CurrentFileName 
    WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Configuring Group Policy '$gpoName': "

    #Office 2013
    
    Set-GPRegistryValue -Name $GpoName -Key "HKCU\Software\Policies\Microsoft\office\15.0\osm" -ValueName CommonFileShare -Type String -Value $CommonFileShare | Out-Null
    Set-GPRegistryValue -Name $GpoName -Key "HKCU\Software\Policies\Microsoft\office\15.0\osm" -ValueName Enablelogging -Type Dword -Value 1 | Out-Null
    Set-GPRegistryValue -Name $GpoName -Key "HKCU\Software\Policies\Microsoft\office\15.0\osm" -ValueName EnableUpload -Type Dword -Value 1 | Out-Null

    #Office 2016
    Set-GPRegistryValue -Name $GpoName -Key "HKCU\Software\Policies\Microsoft\office\16.0\osm" -ValueName CommonFileShare -Type String -Value $CommonFileShare | Out-Null
    Set-GPRegistryValue -Name $GpoName -Key "HKCU\Software\Policies\Microsoft\office\16.0\osm" -ValueName Enablelogging -Type Dword -Value 1 | Out-Null
    Set-GPRegistryValue -Name $GpoName -Key "HKCU\Software\Policies\Microsoft\office\16.0\osm" -ValueName EnableUpload -Type Dword -Value 1 | Out-Null

    Set-TelemetryStartup -GpoName $GpoName -CommonFileShare $CommonFileShare -agentShare $agentShare -ScriptName $ScriptName -TelemetryServer $TelemetryServer

    Write-Host "Done"
    <# write log#>
    $lineNum = Get-CurrentLineNumber    
    $filName = Get-CurrentFileName 
    WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Done"

    Write-Host
    Write-Host "The Group Policy '$gpoName' has been set to configure client to submit telemetry"
    Write-Host
    <# write log#>
    $lineNum = Get-CurrentLineNumber    
    $filName = Get-CurrentFileName 
    WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "The Group Policy '$gpoName' has been set to configure client to submit telemetry"

    if (!($existingGPO)) 
    {
        Write-Host "The Group Policy will not become Active until it linked to an Active Directory Organizational Unit (OU)." `
                   "In Group Policy Management Console link the GPO titled '$gpoName' to the proper OU in your environment." -BackgroundColor Red -ForegroundColor White
                   <# write log#>
    $lineNum = Get-CurrentLineNumber    
    $filName = Get-CurrentFileName 
    WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "The Group Policy will not become Active until it linked to an Active Directory Organizational Unit (OU)."
    WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "In Group Policy Management Console link the GPO titled '$gpoName' to the proper OU in your environment."
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