try {
$enumDef = "
using System;
       [FlagsAttribute]
       public enum Bitness
       {
          Both = 0,
          v32 = 1,
          v64 = 2
       }
"
Add-Type -TypeDefinition $enumDef -ErrorAction SilentlyContinue
} catch { }

try {
$enumDef = "
using System;
       [FlagsAttribute]
       public enum OfficeBranch
       {
          FirstReleaseCurrent = 0,
          Current = 1,
          FirstReleaseBusiness = 2,
          Business = 3,
          CMValidation = 4,
          MonthlyTargeted = 5,
          Monthly = 6,
          SemiAnnualTargeted = 7,
          SemiAnnual = 8
       }
"
Add-Type -TypeDefinition $enumDef -ErrorAction SilentlyContinue
} catch { }

try {
$enumDef = "
using System;
       [FlagsAttribute]
       public enum OfficeChannel
       {
          FirstReleaseCurrent = 0,
          Current = 1,
          FirstReleaseDeferred = 2,
          Deferred = 3,
          MonthlyTargeted = 4,
          Monthly = 5,
          SemiAnnualTargeted = 6,
          SemiAnnual = 7
       }
"
Add-Type -TypeDefinition $enumDef -ErrorAction SilentlyContinue
} catch { }

try {
$enum = "
using System;
 
    [FlagsAttribute]
    public enum GPODeploymentType
    {
        DeployWithScript = 0,
        DeployWithConfigurationFile = 1,
        DeployWithInstallationFile = 2,
        RemoveWithScript = 3
    }
"
Add-Type -TypeDefinition $enum -ErrorAction SilentlyContinue
} catch { }

function Download-GPOOfficeChannelFiles() {
<#
.SYNOPSIS
Downloads the Office Click-to-Run files into the specified folder.

.DESCRIPTION
Downloads the Office 365 ProPlus installation files to a specified file path.

.PARAMETER Channels
The update channel. Current, Deferred, FirstReleaseDeferred, FirstReleaseCurrent

.PARAMETER OfficeFilesPath
This is the location where the source files will be downloaded.

.PARAMETER Languages
All office languages are supported in the ll-cc format "en-us"

.PARAMETER Bitness
Downloads the bitness of Office Click-to-Run "v32, v64, Both"

.PARAMETER Version
You can specify the version to download. 16.0.6868.2062. Version information can be found here https://technet.microsoft.com/en-us/library/mt592918.aspx

.PARAMETER DownloadThrottledVersions
Downloads the version of Office regardless of the throttle value. Set to $false to download non-throttled versions.

.EXAMPLE
Download-GPOOfficeChannelFiles -OfficeFilesPath D:\OfficeChannelFiles

.EXAMPLE
Download-GPOOfficeChannelFiles -OfficeFilesPath D:\OfficeChannelFiles -Channels Deferred -Bitness v32

.EXAMPLE
Download-GPOOfficeChannelFiles -OfficeFilesPath D:\OfficeChannelFiles -Bitness v32 -Channels Deferred,FirstReleaseDeferred -Languages en-us,es-es,ja-jp
#>

    [CmdletBinding(SupportsShouldProcess=$true)]
    Param
    (
        [Parameter()]
        [OfficeChannel[]] $Channels = @(0,1,2,3,4,5,6,7),

        [Parameter(Mandatory=$true)]
	    [String]$OfficeFilesPath = $NULL,

        [Parameter()]
        [ValidateSet("en-us","ar-sa","bg-bg","zh-cn","zh-tw","hr-hr","cs-cz","da-dk","nl-nl","et-ee","fi-fi","fr-fr","de-de","el-gr","he-il","hi-in","hu-hu","id-id","it-it",
                    "ja-jp","kk-kz","ko-kr","lv-lv","lt-lt","ms-my","nb-no","pl-pl","pt-br","pt-pt","ro-ro","ru-ru","sr-latn-rs","sk-sk","sl-si","es-es","sv-se","th-th",
                    "tr-tr","uk-ua","vi-vn")]
        [string[]] $Languages = ("en-us"),

        [Parameter()]
        [Bitness] $Bitness = 0,

        [Parameter()]
        [string] $Version = $NULL,

        [Parameter()]
        [bool] $DownloadThrottledVersions = $true,

        [Parameter()]
        [int] $NumOfRetries = 5,
        
        [Parameter()]
        [bool] $IncludeChannelInfo = $false,

        [Parameter()]
        [bool] $OverWrite = $false,

        [Parameter()]
        [string]$LogFilePath        
    )

    Process {
        $currentFileName = Get-CurrentFileName
        Set-Alias -name LINENUM -value Get-CurrentLineNumber

        if (Test-Path "$PSScriptRoot\Download-OfficeProPlusChannels.ps1") {
           . "$PSScriptRoot\Download-OfficeProPlusChannels.ps1"
        } else {
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 

            WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Dependency file missing: $PSScriptRoot\Download-OfficeProPlusChannels.ps1" -LogFilePath $LogFilePath
           
            throw "Dependency file missing: $PSScriptRoot\Download-OfficeProPlusChannels.ps1"
        }
       
        Download-OfficeProPlusChannels -TargetDirectory $OfficeFilesPath  -Channels $Channels -Version $Version -UseChannelFolderShortName $true -Languages $Languages -Bitness $Bitness `
                                       -DownloadThrottledVersions $DownloadThrottledVersions -NumOfRetries $NumOfRetries -IncludeChannelInfo $IncludeChannelInfo -OverWrite $OverWrite

        $cabFilePath = "$env:TEMP/ofl.cab"
        Copy-Item -Path $cabFilePath -Destination "$OfficeFilesPath\ofl.cab" -Force
    }
}

Function Configure-GPOOfficeDeployment {
<#
.SYNOPSIS
Configures an Office deployment using Group Policy

.DESCRIPTION
Configures the folders and files to deploy Office using Group Policy

.PARAMETER Channel
The update channel to deploy.

.PARAMETER Bitness
The architecture of the update channel.

.PARAMETER OfficeFilesPath
The path to the required deployment files.

.PARAMETER MoveSourceFiles
By default, the installation files will be moved to the source folder. Set this to $false to copy the installation files.

.EXAMPLE
Configure-GPOOfficeDeployment -Channels Current,Deferred,FirstReleaseDeferred -OfficeSourceFilesPath D:\OfficeChannelFiles

.EXAMPLE
Configure-GPOOfficeDeployment -Channels Current,Deferred,FirstReleaseDeferred -OfficeSourceFilesPath D:\OfficeChannelFiles -MoveSourceFiles $false
#>
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param
    (      
        [Parameter()]
        [OfficeChannel[]]$Channels = @(0,1,2,3,4,5,6,7),

        [Parameter()]
        [Bitness]$Bitness = "v32",

        [Parameter()]
	    [string]$OfficeFilesPath,

        [Parameter()]
        [string]$MoveSourceFiles = $true,

        [Parameter()]
        [string]$LogFilePath 
    )

    Begin
    {
	    $currentExecutionPolicy = Get-ExecutionPolicy
	    Set-ExecutionPolicy Unrestricted -Scope Process -Force  
        $startLocation = Get-Location        
    }

    Process 
    {
        Try{
            $currentFileName = Get-CurrentFileName
            Set-Alias -name LINENUM -value Get-CurrentLineNumber

            $cabFilePath = "$OfficeFilesPath\ofl.cab"
            if(Test-Path $cabFilePath){
                Copy-Item -Path $cabFilePath -Destination "$PSScriptRoot\ofl.cab" -Force
            }

            $ChannelList = @("FirstReleaseCurrent","Current","Deferred","FirstReleaseDeferred","MonthlyTargeted","Monthly","SemiAnnualTargeted","SemiAnnual")
            $ChannelXml = Get-ChannelXml -FolderPath $OfficeFilesPath -OverWrite $false
            
            foreach($Channel in $ChannelList){
                if($Channels -contains $Channel){
                    $selectChannel = $ChannelXml.UpdateFiles.baseURL | Where {$_.branch -eq $Channel.ToString() }
                    $latestVersion = Get-ChannelLatestVersion -ChannelUrl $selectChannel.URL -Channel $Channel -FolderPath $OfficeFilesPath -OverWrite $false
        
                    $ChannelShortName = ConvertChannelNameToShortName -ChannelName $Channel
                    $LargeDrv = Get-LargestDrive
        
                    $Path = CreateOfficeChannelShare -Path "$LargeDrv\OfficeDeployment"
        
                    $ChannelPath = "$Path\$Channel"
                    $LocalPath = "$LargeDrv\OfficeDeployment"
                    $LocalChannelPath = "$LargeDrv\OfficeDeployment\SourceFiles"
        
                    [System.IO.Directory]::CreateDirectory($LocalChannelPath) | Out-Null
                           
                    if($OfficeFilesPath) {
                        $officeFileChannelPath = "$OfficeFilesPath\$ChannelShortName"
                        $officeFileTargetPath = "$LocalChannelPath"

                        [string]$oclVersion = $NULL
                        if ($officeFileChannelPath) {
                            if (Test-Path -Path "$officeFileChannelPath\Office\Data") {
                               $oclVersion = Get-LatestVersion -UpdateURLPath $officeFileChannelPath
                            }
                        }

                        if ($oclVersion) {
                           $latestVersion = $oclVersion
                        }

                        if (!(Test-Path -Path $officeFileChannelPath)) {
                            WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Channel Folder Missing: $officeFileChannelPath - Ensure that you have downloaded the Channel you are trying to deploy" -LogFilePath $LogFilePath
                            throw "Channel Folder Missing: $officeFileChannelPath - Ensure that you have downloaded the Channel you are trying to deploy"
                        }

                        [System.IO.Directory]::CreateDirectory($officeFileTargetPath) | Out-Null

                        if ($MoveSourceFiles) {
                            Move-Item -Path $officeFileChannelPath -Destination $officeFileTargetPath -Force
                        } else {
                            Copy-Item -Path $officeFileChannelPath -Destination $officeFileTargetPath -Recurse -Force
                        }

                        $cabFilePath = "$OfficeFilesPath\ofl.cab"
                        if (Test-Path $cabFilePath) {
                            Copy-Item -Path $cabFilePath -Destination "$LocalPath\ofl.cab" -Force
                        }
                    } else {
                        if(Test-Path -Path "$LocalChannelPath\Office") {
                            Remove-Item -Path "$LocalChannelPath\Office" -Force -Recurse
                        }
                    }
        
                    $cabFilePath = "$env:TEMP/ofl.cab"
                    if(!(Test-Path $cabFilePath)) {
                        Copy-Item -Path "$LocalPath\ofl.cab" -Destination $cabFilePath -Force
                    }

                    CreateMainCabFiles -LocalPath $LocalPath -ChannelShortName $ChannelShortName -LatestVersion $latestVersion

                    $DeploymentFilePath = "$PSSCriptRoot\DeploymentFiles\*.*"
                    if (Test-Path -Path $DeploymentFilePath) {
                        Copy-Item -Path $DeploymentFilePath -Destination "$LocalPath" -Force -Recurse
                    } else {
                        WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Deployment folder missing: $DeploymentFilePath" -LogFilePath $LogFilePath
                        throw "Deployment folder missing: $DeploymentFilePath"
                    }
                }
            }
        } Catch{}
    }        
}

Function Create-GPOOfficeDeployment {
<#
.SYNOPSIS
Configures an Office deployment using Group Policy

.DESCRIPTION
Configures an existing Group Policy Object to deploy Office 365 ProPlus 

.PARAMETER GroupPolicyName
The name of the Group Policy Object

.PARAMETER DeploymentType
Choose between DeployWithScript or DeployWithConfigurationFile. DeployWithScript will deploy a dynamic installation
using the target computer's existing Office installation. DeployWithConfigurationFile will deploy a standard Office installation
to all of the targeted computers.

.PARAMETER ScriptName
The name of the deployment script if the DeploymentType is DeployWithScript. If ScriptName is not specified the 
GPO-OfficeDeploymentScript.ps1 will be used.

.PARAMETER OfficeDeploymentFileName
The name of an Office installation file to deploy. An Office install MSI or EXE can be generated using the
Microsoft Office ProPlus Install Toolkit which can be downloaded from http://officedev.github.io/Office-IT-Pro-Deployment-Scripts/XmlEditor.html

.PARAMETER Channel
The update channel to install.

.PARAMETER Bitness
The update channel bit to install.

.PARAMETER ConfigurationXML
The name of a custom (ODT) configuration.xml file if DeploymentTYpe is set to DeployWithConfigurationFile. If you plan on using a custom xml
for the deployment make sure to copy the file to the DeploymentFiles folder before running Configure-GPOOfficeDeployment, or copy the file
to OfficeDeployment if Configure-GPOOfficeDeployment has already been ran.

.PARAMETER WaitForInstallToFinish
While Office is installing PowerShell will remain open until the installation is finished.

.PARAMETER InstallProofingTools
Set this value to $true to include the Proofing Tools exe with the deployment.

.EXAMPLE 
Create-GPOOfficeDeployment -GroupPolicyName DeployCurrentChannel64Bit -DeploymentType DeployWithScript -Channel Current -Bitness 64

.EXAMPLE
Create-GPOOfficeDeployment -GroupPolicyName DeployCurrentChannel32Bit -DeploymentType DeployWithConfigurationFile -Channel Current -Bitness 32 -ConfigurationXML CurrentChannelDeployment.xml

.EXAMPLE
Create-GPOOfficeDeployment -GroupPolicyName DeployWithMSI -DeploymentType DeployWithInstallationFile -OfficeDeploymentFileName OfficeProPlus.msi
#>
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param
    (
	    [Parameter(Mandatory=$True)]
	    [string]$GroupPolicyName,
	
        [Parameter()]
	    [GPODeploymentType]$DeploymentType = 0,
        
        [Parameter()]
        [string]$ScriptName,
              
        [Parameter()]
        [OfficeChannel]$Channel,

        [Parameter()]
        [Bitness]$Bitness,

        [Parameter()]
        [string]$ConfigurationXML = $null,

        [Parameter()]
        [string]$OfficeDeploymentFileName,

        [Parameter()]
        [bool]$WaitForInstallToFinish = $true,

        [Parameter()]
        [bool]$InstallProofingTools = $false,

        [Parameter()]
        [bool]$Quiet = $true,

        [Parameter()]
        [string]$LogFilePath 
    )

    Begin
    {
	    $currentExecutionPolicy = Get-ExecutionPolicy
	    Set-ExecutionPolicy Unrestricted -Scope Process -Force  
        $startLocation = Get-Location        
    }

    Process
    {
        $currentFileName = Get-CurrentFileName
        Set-Alias -name LINENUM -value Get-CurrentLineNumber

        $Root = [ADSI]"LDAP://RootDSE"
        $DomainPath = $Root.Get("DefaultNamingContext")

        Write-Host "Configuring Group Policy to Install Office Click-To-Run"
        Write-Host
        WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Configuring Group Policy to Install Office Click-To-Run" -LogFilePath $LogFilePath

        Write-Host "Searching for GPO: $GroupPolicyName..." -NoNewline
        WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Searching for GPO: $GroupPolicyName..." -LogFilePath $LogFilePath
	    $gpo = Get-GPO -Name $GroupPolicyName
	
	    if(!$gpo -or ($gpo -eq $null))
	    {
            WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "The GPO $GroupPolicyName could not be found." -LogFilePath $LogFilePath
		    Write-Error "The GPO $GroupPolicyName could not be found."
	    }

        Write-Host "GPO Found"
        WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "GPO Found" -LogFilePath $LogFilePath

        Write-Host "Modifying GPO: $GroupPolicyName..." -NoNewline
        WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Modifying GPO: $GroupPolicyName..." -LogFilePath $LogFilePath

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
                      
        $OfficeDeploymentShare = Get-WmiObject Win32_Share | ? {$_.Name -like "OfficeDeployment$"}
        $OfficeDeploymentName = $OfficeDeploymentShare.Name
        $OfficeDeploymentUNC = "\\" + $OfficeDeploymentShare.PSComputerName + "\$OfficeDeploymentName" 
        
        if($Bitness -like "v64"){
            $Bit = "64"
        } else {
            $Bit = "32"
        } 
               
        if($DeploymentType -eq "DeployWithConfigurationFile")
        {
            if(!$ScriptName){$ScriptName = "DeployConfigFile.ps1"}

            $newContent[$nextIndex+1] = "{0}CmdLine={1}" -f $nextScriptIndex, $ScriptName

            if($WaitForInstallToFinish -eq $false){
	            $newContent[$nextIndex+2] = "{0}Parameters=-OfficeDeploymentPath {1} -ConfigFileName {2} -WaitForInstallToFinish {3}" -f $nextScriptIndex, $OfficeDeploymentUNC, $ConfigurationXML, $WaitForInstallToFinish
                if($InstallProofingTools -eq $true){
                    $newContent[$nextIndex+2] = "{0}Parameters=-OfficeDeploymentPath {1} ConfigFileName {2} -WaitForInstallToFinish {3} -InstallProofingTools {4}" -f $nextScriptIndex, $OfficeDeploymentUNC, $ConfigurationXML, $WaitForInstallToFinish, $InstallProofingTools
                }
            } else {
                if($InstallProofingTools -eq $true){
                    $newContent[$nextIndex+2] = "{0}Parameters=-OfficeDeploymentPath {1} -ConfigFileName {2} -InstallProofingTools {3}" -f $nextScriptIndex, $OfficeDeploymentUNC, $ConfigurationXML, $InstallProofingTools
                } else {
                    $newContent[$nextIndex+2] = "{0}Parameters=-OfficeDeploymentPath {1} -ConfigFileName {2}" -f $nextScriptIndex, $OfficeDeploymentUNC, $ConfigurationXML
                }
            }
        } elseif ($DeploymentType -eq "DeployWithScript") 
        {
            if(!$ScriptName){$ScriptName = "GPO-OfficeDeploymentScript.ps1"}

            $newContent[$nextIndex+1] = "{0}CmdLine={1}" -f $nextScriptIndex, $ScriptName

            if($Channel -eq $null -and $Bitness -eq $null){
                $newContent[$nextIndex+2] = "{0}Parameters=-OfficeDeploymentPath {1}" -f $nextScriptIndex, $OfficeDeploymentUNC
            }
            elseif($Channel -eq $null){
                $newContent[$nextIndex+2] = "{0}Parameters=-OfficeDeploymentPath {1} -Bitness {2}" -f $nextScriptIndex, $OfficeDeploymentUNC, $Bit
            }
            elseif($Bitness -eq $null){
                $newContent[$nextIndex+2] = "{0}Parameters=-OfficeDeploymentPath {1} -Channel {2}" -f $nextScriptIndex, $OfficeDeploymentUNC, $Channel
            } else {
                $newContent[$nextIndex+2] = "{0}Parameters=-OfficeDeploymentPath {1} -Channel {2} -Bitness {3}" -f $nextScriptIndex, $OfficeDeploymentUNC, $Channel, $Bit
            }

        } elseif($DeploymentType -eq "DeployWithInstallationFile")
        {
            if(!$ScriptName){$ScriptName = "DeployOfficeInstallationFile.ps1"}
            if(!$OfficeDeploymentFileName){$OfficeDeploymentFileName = "OfficeProPlus.msi"}
            
            $Quiet = Convert-Bool $Quiet
            
            $newContent[$nextIndex+1] = "{0}CmdLine={1}" -f $nextScriptIndex, $ScriptName
            $newContent[$nextIndex+2] = "{0}Parameters=-OfficeDeploymentPath {1} -OfficeDeploymentFileName {2} -Quiet {3}" -f $nextScriptIndex, $OfficeDeploymentUNC, $OfficeDeploymentFileName, $Quiet

        } elseif($DeploymentType -eq "RemoveWithScript")
        {
            if(!$ScriptName){$ScriptName = "GPO-ExampleRemovePreviousOfficeInstalls.ps1"}
            
            $newContent[$nextIndex+1] = "{0}CmdLine={1}" -f $nextScriptIndex, $ScriptName
            $newContent[$nextIndex+2] = "{0}Parameters=-OfficeDeploymentPath {1}" -f $nextScriptIndex, $OfficeDeploymentUNC
        }

	    for($i=$nextIndex; $i -lt $length; $i++)
	    {
		    $newContent[$i] = $content[$i]
	    }

	    $newContent | Set-Content -Encoding $encoding -Path $psScriptsFilePath -Force
	    #endregion
	
	    #region Place the script to attach in the StartUp Folder
        $LargeDrv = Get-LargestDrive 
	    $setupExeSourcePath = "$LargeDrv\OfficeDeployment\$ScriptName"
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
            $extSplit = $currentExt.Split(']')

            foreach ($extGuid in $extSplit) {
              if ($extGuid) {
                if ($extGuid.Length -gt 0) {
                    $addItem = $extList.Add($extGuid.Replace("[", "").ToUpper())
                }
              }
            }
        }

        $extGuids = @("{42B5FAAE-6536-11D2-AE5A-0000F87571E3}{40B6664F-4972-11D1-A7CA-0000F87571E3}")

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
        Write-Host "The Group Policy '$GroupPolicyName' has been modified to install Office at Workstation Startup." -BackgroundColor DarkBlue
        Write-Host "Once Group Policy has refreshed on the Workstations then Office will install on next startup if the computer has access to the Network Share." -BackgroundColor DarkBlue

        WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "GPO Modified" -LogFilePath $LogFilePath
        WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "The Group Policy '$GroupPolicyName' has been modified to install Office at Workstation Startup." -LogFilePath $LogFilePath
        WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Once Group Policy has refreshed on the Workstations then Office will install on next startup if the computer has access to the Network Share." -LogFilePath $LogFilePath
    }

    End 
    {      
       $setLocation = Set-Location $startLocation
    }
}

Function Update-GPOSourceFiles {
<#
.SYNOPSIS
Updates the SourceFiles of a GPO deployment

.DESCRIPTION
Update the SourceFiles by copying or moving additional channel files to the OfficeDeployment$ share.

.PARAMETER OfficeFilesPath
The filepath to where the Office channel files are downloaded.

.PARAMETER Channels
The channels to copy/move to the OfficeDeployment$ share.

.PARAMETER MoveSourceFiles
If set to $true the channel files will be moved to the OfficeDeployment$ share. If not specified the files will be copied.

.EXAMPLE
Update-GPOSourceFiles -OfficeFilesPath D:\OfficeChannelFiles -Channel Current

.EXAMPLE
Update-GPOSourceFiles -OfficeFilesPath D:\OfficeChannelFiles -Channels Deferred,FirstReleaseDeferred -MoveSourceFiles $true
#>
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param(
        [Parameter()]
	    [string]$OfficeFilesPath,

        [Parameter()]
        [OfficeChannel[]] $Channels = @(0,1,2,3,4,5,6,7),

        [Parameter()]
	    [bool]$MoveSourceFiles = $false
    )

    Begin
    {
	    $currentExecutionPolicy = Get-ExecutionPolicy
	    Set-ExecutionPolicy Unrestricted -Scope Process -Force  
        $startLocation = Get-Location        
    }

    Process{
        try{
            Check-AdminAccess

            $cabFilePath = "$OfficeFilesPath\ofl.cab"

            if (Test-Path $cabFilePath) {
                Copy-Item -Path $cabFilePath -Destination "$PSScriptRoot\ofl.cab" -Force
            }

            $OfficeDeploymentShare = Get-WmiObject Win32_Share | ? {$_.Name -like "OfficeDeployment$"}
            $OfficeDeploymentName = $OfficeDeploymentShare.Name
            $OfficeDeploymentUNC = "\\" + $OfficeDeploymentShare.PSComputerName + "\$OfficeDeploymentName"
            $LocalChannelPath = $OfficeDeploymentUNC 

            foreach($Channel in $Channels){
                $ChannelShortName = ConvertChannelNameToShortName -ChannelName $Channel

                $officeFileChannelPath = "$OfficeFilesPath\$ChannelShortName"
                $officeFileTargetPath = "$LocalChannelPath\SourceFiles\$ChannelShortName"

                $tempofficeFileChannelPath = "$officeFileChannelPath\Office\Data"
                $tempLocalChannelPath = "$LocalChannelPath\SourceFiles\$ChannelShortName\Office\Data"

                [string]$oclVersion = $NULL
                if ($officeFileChannelPath) {
                    if (Test-Path -Path "$officeFileChannelPath\Office\Data") {
                       $oclVersion = Get-LatestVersion -UpdateURLPath $officeFileChannelPath
                    }
                }

                if ($oclVersion) {
                    $latestVersion = $oclVersion
                }

                if ($MoveSourceFiles){                                
                    if(!(Test-Path -Path $officeFileTargetPath)) {
                        Move-Item -Path $officeFileChannelPath -Destination "$LocalChannelPath\SourceFiles" -Force
                    }else{
                        $subfiles = Get-ChildItem $tempofficeFileChannelPath
                        foreach($file in $subfiles){
                            [array]$tempLocalChannelPathFiles = (Get-ChildItem -Path $tempLocalChannelPath).Name
                            if($tempLocalChannelPathFiles -notcontains $file.Name){
                                Move-Item -Path $tempofficeFileChannelPath\$file -Destination $tempLocalChannelPath -Force
                            }
                            else{
                                [array]$versionFiles = (Get-ChildItem -Path $tempLocalChannelPath\$latestVersion).Name
                                [array]$officeChannelVersionFiles = (Get-ChildItem -Path "$tempofficeFileChannelPath\$latestVersion").Name
                                foreach($officeChannelVersionFile in $officeChannelVersionFiles) {
                                    if($versionFiles -notcontains $officeChannelVersionFile){
                                        Move-Item -Path $tempofficeFileChannelPath\$latestVersion\$officeChannelVersionFile -Destination $tempLocalChannelPath\$latestVersion -Force
                                    }
                                }
                            }           
                        }

                        Get-ChildItem -Path $officeFileChannelPath -Recurse | Remove-Item -Force -Recurse | Out-Null

                        [System.IO.Directory]::Delete($officeFileChannelPath) | Out-Null
                    }
                } else {
                    if(!(Test-Path -Path $officeFileTargetPath)) {
                        #[System.IO.Directory]::CreateDirectory($officeFileTargetPath) | Out-Null

                        Copy-Item -Path $officeFileChannelPath -Destination "$LocalChannelPath\SourceFiles" -Recurse -Force
                    }
                    else{
                        $subfiles = Get-ChildItem $tempofficeFileChannelPath
                        foreach($file in $subfiles){
                            Copy-Item -Path $tempofficeFileChannelPath\$file -Destination $tempLocalChannelPath -Recurse -Force 
                        }             
                    }
                }

                $cabFilePath = "$OfficeFilesPath\ofl.cab"
                if (Test-Path $cabFilePath) {
                    Copy-Item -Path $cabFilePath -Destination "$LocalPath\ofl.cab" -Force
                }

                CreateMainCabFiles -LocalPath $LocalChannelPath -ChannelShortName $ChannelShortName -LatestVersion $latestVersion
            }
        } catch {}
    }
}


function ConvertChannelNameToShortName {
    Param(
       [Parameter()]
       [string] $ChannelName
    )
    Process {
       if ($ChannelName.ToLower() -eq "FirstReleaseCurrent".ToLower()) {
         return "FRCC"
       }
       if ($ChannelName.ToLower() -eq "Current".ToLower()) {
         return "CC"
       }
       if ($ChannelName.ToLower() -eq "FirstReleaseDeferred".ToLower()) {
         return "FRDC"
       }
       if ($ChannelName.ToLower() -eq "Deferred".ToLower()) {
         return "DC"
       }
       if ($ChannelName.ToLower() -eq "Business".ToLower()) {
         return "DC"
       }
       if ($ChannelName.ToLower() -eq "FirstReleaseBusiness".ToLower()) {
         return "FRDC"
       }
       if ($ChannelName.ToLower() -eq "MonthlyTargeted".ToLower()) {
         return "MTC"
       }
       if ($ChannelName.ToLower() -eq "Monthly".ToLower()) {
         return "MC"
       }
       if ($ChannelName.ToLower() -eq "SemiAnnualTargeted".ToLower()) {
         return "SATC"
       }
       if ($ChannelName.ToLower() -eq "SemiAnnual".ToLower()) {
         return "SAC"
       }
    }
}

function ConvertChannelNameToBranchName {
    Param(
       [Parameter()]
       [string] $ChannelName
    )
    Process {
       if ($ChannelName.ToLower() -eq "FirstReleaseCurrent".ToLower()) {
         return "FirstReleaseCurrent"
       }
       if ($ChannelName.ToLower() -eq "Current".ToLower()) {
         return "Current"
       }
       if ($ChannelName.ToLower() -eq "FirstReleaseDeferred".ToLower()) {
         return "FirstReleaseBusiness"
       }
       if ($ChannelName.ToLower() -eq "Deferred".ToLower()) {
         return "Business"
       }
       if ($ChannelName.ToLower() -eq "Business".ToLower()) {
         return "Business"
       }
       if ($ChannelName.ToLower() -eq "FirstReleaseBusiness".ToLower()) {
         return "FirstReleaseBusiness"
       }
       if ($ChannelName.ToLower() -eq "MonthlyTargeted".ToLower()) {
         return "MonthlyTargeted"
       }
       if ($ChannelName.ToLower() -eq "Monthly".ToLower()) {
         return "Monthly"
       }
       if ($ChannelName.ToLower() -eq "SemiAnnualTargeted".ToLower()) {
         return "SemiAnnualTargeted"
       }
       if ($ChannelName.ToLower() -eq "SemiAnnual".ToLower()) {
         return "SemiAnnual"
       }
    }
}

function ConvertBranchNameToChannelName {
    Param(
       [Parameter()]
       [string] $BranchName
    )
    Process {
       if ($BranchName.ToLower() -eq "FirstReleaseCurrent".ToLower()) {
         return "FirstReleaseCurrent"
       }
       if ($BranchName.ToLower() -eq "Current".ToLower()) {
         return "Current"
       }
       if ($BranchName.ToLower() -eq "FirstReleaseDeferred".ToLower()) {
         return "FirstReleaseDeferred"
       }
       if ($BranchName.ToLower() -eq "Deferred".ToLower()) {
         return "Deferred"
       }
       if ($BranchName.ToLower() -eq "Business".ToLower()) {
         return "Deferred"
       }
       if ($BranchName.ToLower() -eq "FirstReleaseBusiness".ToLower()) {
         return "FirstReleaseDeferred"
       }
       if ($BranchName.ToLower() -eq "MonthlyTargeted".ToLower()) {
         return "MonthlyTargeted"
       }
       if ($BranchName.ToLower() -eq "Monthly".ToLower()) {
         return "Monthly"
       }
       if ($BranchName.ToLower() -eq "SemiAnnualTargeted".ToLower()) {
         return "SemiAnnualTargeted"
       }
       if ($BranchName.ToLower() -eq "SemiAnnual".ToLower()) {
         return "SemiAnnual"
       }
    }
}

function CreateMainCabFiles() {
    [CmdletBinding()]	
    Param
	(
		[Parameter(Mandatory=$true)]
		[String]$LocalPath = "",

        [Parameter(Mandatory=$true)]
        [String] $ChannelShortName,

        [Parameter(Mandatory=$true)]
        [String] $LatestVersion
	) 
    Process {
        $versionFile321 = "$LocalPath\$ChannelShortName\Office\Data\v32_$LatestVersion.cab"
        $v32File1 = "$LocalPath\$ChannelShortName\Office\Data\v32.cab"

        $versionFile641 = "$LocalPath\$ChannelShortName\Office\Data\v64_$LatestVersion.cab"
        $v64File1 = "$LocalPath\$ChannelShortName\Office\Data\v64.cab"

        $versionFile322 = "$LocalPath\SourceFiles\$ChannelShortName\Office\Data\v32_$LatestVersion.cab"
        $v32File2 = "$LocalPath\SourceFiles\$ChannelShortName\Office\Data\v32.cab"

        $versionFile642 = "$LocalPath\SourceFiles\$ChannelShortName\Office\Data\v64_$LatestVersion.cab"
        $v64File2 = "$LocalPath\SourceFiles\$ChannelShortName\Office\Data\v64.cab"

        if (Test-Path -Path $versionFile321) {
            Copy-Item -Path $versionFile321 -Destination $v32File1 -Force
        }

        if (Test-Path -Path $versionFile641) {
            Copy-Item -Path $versionFile641 -Destination $v64File1 -Force
        }

        if (Test-Path -Path $versionFile322) {
            Copy-Item -Path $versionFile322 -Destination $v32File2 -Force
        }

        if (Test-Path -Path $versionFile642) {
            Copy-Item -Path $versionFile642 -Destination $v64File2 -Force
        }
    }
}

function CreateOfficeChannelShare() {
    [CmdletBinding()]	
    Param
	(
        [Parameter()]
        [String]$Name = "OfficeDeployment$",

        [Parameter()]
        [String]$Path = "$env:SystemDrive\OfficeDeployment"
	) 
    
    if (!(Test-Path $Path)) { 
      $addFolder = New-Item $Path -type Directory 
    }
    
    $ACL = Get-ACL $Path

    $identity = New-Object System.Security.Principal.NTAccount  -argumentlist ("$env:UserDomain\$env:UserName") 
    $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule -argumentlist ($identity,"FullControl","ContainerInherit, ObjectInherit","None","Allow")

    $addAcl = $ACL.AddAccessRule($accessRule) | Out-Null

    $identity = New-Object System.Security.Principal.NTAccount -argumentlist ("$env:UserDomain\Domain Admins") 
    $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule -argumentlist ($identity,"FullControl","ContainerInherit, ObjectInherit","None","Allow")
    $addAcl = $ACL.AddAccessRule($accessRule) | Out-Null

    $identity = "Everyone"
    $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule -argumentlist ($identity,"Read","ContainerInherit, ObjectInherit","None","Allow")
    $addAcl = $ACL.AddAccessRule($accessRule) | Out-Null

    Set-ACL -Path $Path -ACLObject $ACL | Out-Null
    
    $share = Get-WmiObject -Class Win32_share | Where {$_.name -eq "$Name"}
    if (!$share) {
       Create-FileShare -Name $Name -Path $Path | Out-Null
    }

    $sharePath = "\\$env:COMPUTERNAME\$Name"
    return $sharePath
}

function Get-ChannelLatestVersion() {
   [CmdletBinding()]
   param( 
      [Parameter(Mandatory=$true)]
      [string]$ChannelUrl,

      [Parameter(Mandatory=$true)]
      [string]$Channel,

	  [Parameter()]
	  [string]$FolderPath = $null,

	  [Parameter()]
	  [bool]$OverWrite = $false
   )

   process {

       [bool]$downloadFile = $true

       $channelShortName = ConvertChannelNameToShortName -ChannelName $Channel

       if (!($OverWrite)) {
          if ($FolderPath) {
              $CABFilePath = "$FolderPath\$channelShortName\Office\Data\v32.cab"

              if (!(Test-Path -Path $CABFilePath)) {
                 $CABFilePath = "$FolderPath\$channelShortName\Office\Data\v64.cab"
              }

              if (Test-Path -Path $CABFilePath) {
                 $downloadFile = $false
              } else {
                throw "File missing $FolderPath\$channelShortName\Office\Data\v64.cab or $FolderPath\$channelShortName\Office\Data\v64.cab"
              }
          }
       }

       if ($downloadFile) {
           $webclient = New-Object System.Net.WebClient
           $CABFilePath = "$env:TEMP/v32.cab"
           $XMLDownloadURL = "$ChannelUrl/Office/Data/v32.cab"
           $webclient.DownloadFile($XMLDownloadURL,$CABFilePath)

           if ($FolderPath) {
             [System.IO.Directory]::CreateDirectory($FolderPath) | Out-Null

             $channelShortName = ConvertChannelNameToShortName -ChannelName $Channel 

             $targetFile = "$FolderPath\$channelShortName\v32.cab"
             Copy-Item -Path $CABFilePath -Destination $targetFile -Force
           }
       }

       $tmpName = "VersionDescriptor.xml"
       expand $CABFilePath $env:TEMP -f:$tmpName | Out-Null
       $tmpName = $env:TEMP + "\VersionDescriptor.xml"
       [xml]$versionXml = Get-Content $tmpName

       return $versionXml.Version.Available.Build
   }
}

function Get-ChannelXml() {
    [CmdletBinding()]	
    Param
	(
	    [Parameter()]
	    [string]$FolderPath = $null,

	    [Parameter()]
	    [bool]$OverWrite = $false,

        [Parameter()]
        [string] $Bitness = "32"
	)

   process {
       $cabPath = "$PSScriptRoot\ofl.cab"
       [bool]$downloadFile = $true

       if (!($OverWrite)) {
          if ($FolderPath) {
              $XMLFilePath = "$FolderPath\ofl.cab"
              if (Test-Path -Path $XMLFilePath) {
                 $downloadFile = $false
              } else {
                throw "File missing $FolderPath\ofl.cab"
              }
          }
       }

       if ($downloadFile) {
           $webclient = New-Object System.Net.WebClient
           $XMLFilePath = "$env:TEMP/ofl.cab"
           $XMLDownloadURL = "http://officecdn.microsoft.com/pr/wsus/ofl.cab"
           $webclient.DownloadFile($XMLDownloadURL,$XMLFilePath)

           if ($FolderPath) {
             [System.IO.Directory]::CreateDirectory($FolderPath) | Out-Null
             $targetFile = "$FolderPath\ofl.cab"
             Copy-Item -Path $XMLFilePath -Destination $targetFile -Force
           }
       }

       $tmpName = "o365client_" + $Bitness + "bit.xml"
       expand $XMLFilePath $env:TEMP -f:$tmpName | Out-Null
       $tmpName = $env:TEMP + "\" + $tmpName
       
       [xml]$channelXml = Get-Content $tmpName

       return $channelXml
   }

}

function Get-LargestDrive() {
   [CmdletBinding()]
   param( 
   )
   process {
      $drives = Get-Partition | where {$_.DriveLetter}
      $driveInfoList = @()

      foreach ($drive in $drives) {
          $driveLetter = $drive.DriveLetter
          $deviceFilter = "DeviceID='" + $driveLetter + ":'" 
 
          $driveInfo = Get-WmiObject Win32_LogicalDisk -ComputerName "." -Filter $deviceFilter
          $driveInfoList += $driveInfo
      }

      $itemList = @()
      foreach($item in $driveInfoList){
          $itemList += $item.Freespace
      }

      $largItem = $itemList | measure -Maximum      
      $largestItem = $largItem.Maximum

      $FreeSpaceDrive = $driveInfoList | Where-Object {$_.Freespace -eq $largestItem}

      return $FreeSpaceDrive.DeviceID
   }
}

Function Get-LatestVersion() {
  [CmdletBinding()]
  Param(
     [Parameter(Mandatory=$true)]
     [string] $UpdateURLPath
  )

  process {
    [array]$totalVersion = @()
    $Version = $null

    $LatestBranchVersionPath = $UpdateURLPath + '\Office\Data'
    if(Test-Path $LatestBranchVersionPath){
        $DirectoryList = Get-ChildItem $LatestBranchVersionPath
        Foreach($listItem in $DirectoryList){
            if($listItem.GetType().Name -eq 'DirectoryInfo'){
                $totalVersion+=$listItem.Name
            }
        }
    }

    $totalVersion = $totalVersion | Sort-Object -Descending
    
    #sets version number to the newest version in directory for channel if version is not set by user in argument  
    if($totalVersion.Count -gt 0){
        $Version = $totalVersion[0]
    }

    return $Version
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

function Get-CurrentLineNumber {
    $MyInvocation.ScriptLineNumber
}

function Get-CurrentFileName{
    $MyInvocation.ScriptName.Substring($MyInvocation.ScriptName.LastIndexOf("\")+1)
}

Function WriteToLogFile() {
    param( 
        [Parameter(Mandatory=$true)]
        [string]$LNumber,

        [Parameter(Mandatory=$true)]
        [string]$FName,

        [Parameter(Mandatory=$true)]
        [string]$ActionError,

        [Parameter()]
        [string]$LogFilePath
    )

    try{
        $headerString = "Time".PadRight(30, ' ') + "Line Number".PadRight(15,' ') + "FileName".PadRight(60,' ') + "Action"
        $stringToWrite = $(Get-Date -Format G).PadRight(30, ' ') + $($LNumber).PadRight(15, ' ') + $($FName).PadRight(60,' ') + $ActionError

        if(!$LogFilePath){
            $LogFilePath = "$env:windir\Temp\" + (Get-Date -Format u).Substring(0,10)+"_OfficeDeploymentLog.txt"
        }
        if(Test-Path $LogFilePath){
             Add-Content $LogFilePath $stringToWrite
        }
        else{#if not exists, create new
             Add-Content $LogFilePath $headerString
             Add-Content $LogFilePath $stringToWrite
        }
    } catch [Exception]{
        Write-Host $_
    }
}

Function Convert-Bool() {
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param
    (
        [Parameter(Mandatory=$true)]
        [bool] $value
    )

    $newValue = "$" + $value.ToString()
    return $newValue 
}

$scriptPath = GetScriptRoot

$shareFunctionsPath = "$scriptPath\SharedFunctions.ps1"
if ($scriptPath.StartsWith("\\")) {
} else {
    if (!(Test-Path -Path $shareFunctionsPath)) {
    <# write log#>
        $lineNum = Get-CurrentLineNumber    
        $filName = Get-CurrentFileName 
        WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Missing Dependency File SharedFunctions.ps1"    
        throw "Missing Dependency File SharedFunctions.ps1"    
    }
}
. $shareFunctionsPath