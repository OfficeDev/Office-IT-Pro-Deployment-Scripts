try {
$enum = "
using System;
 
    [FlagsAttribute]
    public enum CMDeploymentType
    {
        DeployWithScript = 0,
        DeployWithConfigurationFile = 1
    }
"
Add-Type -TypeDefinition $enum -ErrorAction SilentlyContinue
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
$enum2 = "
using System;
 
    [FlagsAttribute]
    public enum CMOfficeProgramType
    {
        DeployWithScript = 0,
        DeployWithConfigurationFile = 1,
        ChangeChannel = 2,
        RollBack = 3,
        UpdateWithConfigMgr = 4,
        UpdateWithTask = 5,
        LanguagePack = 6
    }
"
Add-Type -TypeDefinition $enum2 -ErrorAction SilentlyContinue
} catch { }

try {
$enumBitness = "
using System;
       [FlagsAttribute]
       public enum Bitness
       {
          Both = 0,
          v32 = 1,
          v64 = 2
       }
"
Add-Type -TypeDefinition $enumBitness -ErrorAction SilentlyContinue
} catch { }

try {
$enumBitnessOptions = "
using System;
       [FlagsAttribute]
       public enum BitnessOptions
       {
          v32 = 1,
          v64 = 2
       }
"
Add-Type -TypeDefinition $enumBitnessOptions -ErrorAction SilentlyContinue
} catch { }

try {
$deploymentPurpose = "
using System;
       [FlagsAttribute]
       public enum DeploymentPurpose
       {
          Default= 0,
          Required = 1,
          Available = 2
       }
"
Add-Type -TypeDefinition $deploymentPurpose -ErrorAction SilentlyContinue
} catch { }

function Download-CMOfficeChannelFiles() {
<#
.SYNOPSIS
Downloads the Office Click-to-Run files into the specified folder for package creation.

.DESCRIPTION
Downloads the Office 365 ProPlus installation files to a specified file path.

.PARAMETER Channels
The update channel. Current, Deferred, FirstReleaseDeferred, FirstReleaseCurrent

.PARAMETER OfficeFilesPath
This is the location where the source files will be downloaded to

.PARAMETER Languages
All office languages are supported in the ll-cc format "en-us"

.PARAMETER Bitness
Downloads the bitness of Office Click-to-Run "v32, v64, Both"

.PARAMETER Version
You can specify the version to download. 16.0.6868.2062. Version information can be found here https://technet.microsoft.com/en-us/library/mt592918.aspx

.EXAMPLE
Download-CMOfficeChannelFiles -OfficeFilesPath D:\OfficeChannelFiles

.EXAMPLE
Download-CMOfficeChannelFiles -OfficeFilesPath D:\OfficeChannelFiles -Channels Deferred -Bitness v32

.EXAMPLE
Download-CMOfficeChannelFiles -OfficeFilesPath D:\OfficeChannelFiles -Bitness v32 -Channels Deferred,FirstReleaseDeferred -Languages en-us,es-es,ja-jp
#>

    [CmdletBinding(SupportsShouldProcess=$true)]
    Param
    (
        [Parameter()]
        [OfficeChannel[]] $Channels = @(0,1,2,3,4,5,6,7,8,9),

        [Parameter(Mandatory=$true)]
	    [String]$OfficeFilesPath = $NULL,

        [Parameter()]
        [ValidateSet("en-us","MatchOS","ar-sa","bg-bg","zh-cn","zh-tw","hr-hr","cs-cz","da-dk","nl-nl","et-ee","fi-fi","fr-fr","de-de","el-gr","he-il","hi-in","hu-hu","id-id","it-it",
                    "ja-jp","kk-kz","ko-kr","lv-lv","lt-lt","ms-my","nb-no","pl-pl","pt-br","pt-pt","ro-ro","ru-ru","sr-latn-rs","sk-sk","sl-si","es-es","sv-se","th-th",
                    "tr-tr","uk-ua","vi-vn")]
        [string[]] $Languages = ("en-us"),

        [Parameter()]
        [ValidateSet("af-za","sq-al","am-et","hy-am","as-in","az-latn-az","eu-es","be-by","bn-bd","bn-in","bs-latn-ba","ca-es","prs-af","fil-ph","gl-es","ka-ge","gu-in","is-is","ga-ie","kn-in",
                "km-kh","sw-ke","kok-in","ky-kg","lb-lu","mk-mk","ml-in","mt-mt","mi-nz","mr-in","mn-mn","ne-np","nn-no","or-in","fa-ir","pa-in","quz-pe","gd-gb","sr-cyrl-rs","sr-cyrl-ba",
                "sd-arab-pk","si-lk","ta-in","tt-ru","te-in","tk-tm","ur-pk","ug-cn","uz-latn-uz","ca-es-valencia","cy-gb","none")]
        [string[]] $PartialLanguages = ("none"),

        [Parameter()]
        [ValidateSet("ha-latn-ng","ig-ng","xh-za","zu-za","rw-rw","ps-af","rm-ch","nso-za","tn-za","wo-sn","yo-ng","none")]
        [string[]] $ProofingLanguages = ("none"),

        [Parameter()]
        [Bitness] $Bitness = 0,

        [Parameter()]
        [string] $Version = $NULL,

        [Parameter()]
        [string]$LogFilePath
        
    )

    Process {
       $currentFileName = Get-CurrentFileName
       Set-Alias -name LINENUM -value Get-CurrentLineNumber

       if (Test-Path "$PSScriptRoot\Download-OfficeProPlusChannels.ps1") {
         . "$PSScriptRoot\Download-OfficeProPlusChannels.ps1"
       } else {
         WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Dependency file missing: $PSScriptRoot\Download-OfficeProPlusChannels.ps1" -LogFilePath $LogFilePath
         throw "Dependency file missing: $PSScriptRoot\Download-OfficeProPlusChannels.ps1"
       }

       $ChannelList = @("FirstReleaseCurrent","Current","FirstReleaseDeferred","Deferred","MonthlyTargeted","Monthly","SemiAnnualTargeted","SemiAnnual")
       $ChannelXml = Get-ChannelXml -FolderPath $OfficeFilesPath -OverWrite $true

       foreach ($Channel in $ChannelList) {
         if ($Channels -contains $Channel) {

            switch($Channel){
                "FirstReleaseCurrent"{
                    $cabChannelName = ""
                }
                "Current"{
                    $cabChannelName = "Monthly"
                }
                "FirstReleaseDeferred"{
                    $cabChannelName = "Targeted"
                }
                "Deferred"{
                    $cabChannelName = "Broad"
                }
                "MonthlyTargeted"{
                    $cabChannelName = "Insiders"
                }
                "Monthly"{
                    $cabChannelName = "Monthly"
                }
                "SemiAnnualTargeted"{
                    $cabChannelName = "Targeted"
                }
                "SemiAnnual"{
                    $cabChannelName = "Broad"
                }
            }

            $selectChannel = $ChannelXml.UpdateFiles.baseURL | Where {$_.branch -eq $cabChannelName.ToString() }
            $latestVersion = Get-ChannelLatestVersion -ChannelUrl $selectChannel.URL -Channel $cabChannelName
            $ChannelShortName = ConvertChannelNameToShortName -ChannelName $Channel

            if ($Version) {
               $latestVersion = $Version
            }

            Download-OfficeProPlusChannels -TargetDirectory $OfficeFilesPath  -Channels $Channel -Version $latestVersion -UseChannelFolderShortName $true -Languages $Languages -Bitness $Bitness -PartialLanguages $PartialLanguages -ProofingLanguages $ProofingLanguages

            $cabFilePath = "$env:TEMP/ofl.cab"
            WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Copying $cabFilePath to $OfficeFilesPath\ofl.cab" -LogFilePath $LogFilePath
            Copy-Item -Path $cabFilePath -Destination "$OfficeFilesPath\ofl.cab" -Force

            $latestVersion = Get-ChannelLatestVersion -ChannelUrl $selectChannel.URL -Channel $Channel -FolderPath $OfficeFilesPath -OverWrite $true 
         }
       }
    }
}

function Create-CMOfficePackage {
<#

.SYNOPSIS
Automates the configuration of System Center Configuration Manager (CM) to create an Office Click-To-Run Package

.PARAMETER Channels
The update channel. Current, Deferred, FirstReleaseDeferred, FirstReleaseCurrent

.PARAMETER Bitness
Downloads the bitness of Office Click-to-Run "v32, v64, Both"

.PARAMETER OfficeSourceFilesPath
This is the location where the source files are available at

.PARAMETER MoveSourceFiles
This moves the files from the Source location to the location specified

.PARAMETER CustomPackageShareName
This sets a custom package share to use

.PARAMETER UpdateOnlyChangedBits

.PARAMETER SiteCode
The site code you would like to create the package on. If left blank it will default to the current site

.PARAMETER CMPSModulePath
Allows the user to specify that full path to the ConfigurationManager.psd1 PowerShell Module. This is especially useful if CM is installed in a non standard path.

.EXAMPLE
Create-CMOfficePackage -Channels Deferred -Bitness v32 -OfficeSourceFilesPath D:\OfficeChannelFiles

#>
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param
    (
        [Parameter()]
        [OfficeChannel[]] $Channels = @(0,1,2,3,4,5,6,7),

        [Parameter()]
	    [Bitness]$Bitness = "v32",

        [Parameter()]
	    [String]$OfficeSourceFilesPath = $NULL,

        [Parameter()]
	    [bool]$MoveSourceFiles = $false,

		[Parameter()]
		[String]$CustomPackageShareName = $null,

	    [Parameter()]	
	    [Bool]$UpdateOnlyChangedBits = $true,

	    [Parameter()]
	    [String]$SiteCode = $null,

	    [Parameter()]
	    [String]$CMPSModulePath = $NULL,

        [Parameter()]
        [string]$LogFilePath
    )
    Begin
    {
        $currentExecutionPolicy = Get-ExecutionPolicy
	    Set-ExecutionPolicy Unrestricted -Scope Process -Force  
        $startLocation = Get-Location
    }
    Process {
       try {
       $currentFileName = Get-CurrentFileName
       Set-Alias -name LINENUM -value Get-CurrentLineNumber

       Check-AdminAccess

       $cabFilePath = "$OfficeSourceFilesPath\ofl.cab"
       if (Test-Path $cabFilePath) {
            WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Copying $cabFilePath to $PSScriptRoot\ofl.cab" -LogFilePath $LogFilePath
            Copy-Item -Path $cabFilePath -Destination "$PSScriptRoot\ofl.cab" -Force
       }

       $ChannelList = @("FirstReleaseCurrent", "Current", "FirstReleaseDeferred", "Deferred","MonthlyTargeted","Monthly","SemiAnnualTargeted","SemiAnnual")
       $ChannelXml = Get-ChannelXml -FolderPath $OfficeSourceFilesPath -OverWrite $false

       [bool]$packageCreated = $false

       foreach ($Channel in $ChannelList) {
         if ($Channels -contains $Channel) {
         
           switch($Channel){
               "FirstReleaseCurrent"{
                   $cabChannelName = ""
               }
               "Current"{
                   $cabChannelName = "Monthly"
               }
               "FirstReleaseDeferred"{
                   $cabChannelName = "Targeted"
               }
               "Deferred"{
                   $cabChannelName = "Broad"
               }
               "MonthlyTargeted"{
                   $cabChannelName = "Insiders"
               }
               "Monthly"{
                   $cabChannelName = "Monthly"
               }
               "SemiAnnualTargeted"{
                   $cabChannelName = "Targeted"
               }
               "SemiAnnual"{
                   $cabChannelName = "Broad"
               }
           }

           $selectChannel = $ChannelXml.UpdateFiles.baseURL | Where {$_.branch -eq $cabChannelName.ToString() }
           $latestVersion = Get-ChannelLatestVersion -ChannelUrl $selectChannel.URL -Channel $cabChannelName -FolderPath $OfficeFilesPath -OverWrite $false

           $ChannelShortName = ConvertChannelNameToShortName -ChannelName $Channel
           $existingPackage = CheckIfPackageExists
           $LargeDrv = Get-LargestDrive

           $Path = CreateOfficeChannelShare -Path "$LargeDrv\OfficeDeployment"

           $packageName = "Office 365 ProPlus"
           $ChannelPath = "$Path\$Channel"
           $LocalPath = "$LargeDrv\OfficeDeployment"
           $LocalChannelPath = "$LargeDrv\OfficeDeployment\SourceFiles"

           [System.IO.Directory]::CreateDirectory($LocalChannelPath) | Out-Null
                          
           if ($OfficeSourceFilesPath) {
                $officeFileChannelPath = "$OfficeSourceFilesPath\$ChannelShortName"
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
                    WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Moving $officeFileChannelPath to $officeFileTargetPath" -LogFilePath $LogFilePath
                    Move-Item -Path $officeFileChannelPath -Destination $officeFileTargetPath -Force
                } else {
                    WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Copying $officeFileChannelPath to $officeFileTargetPath" -LogFilePath $LogFilePath
                    Copy-Item -Path $officeFileChannelPath -Destination $officeFileTargetPath -Recurse -Force
                }

                $cabFilePath = "$OfficeSourceFilesPath\ofl.cab"
                if (Test-Path $cabFilePath) {
                    WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Copying $cabFilePath to $LocalPath\ofl.cab" -LogFilePath $LogFilePath
                    Copy-Item -Path $cabFilePath -Destination "$LocalPath\ofl.cab" -Force
                }
           } else {
              if (Test-Path -Path "$LocalChannelPath\Office") {
                 Remove-Item -Path "$LocalChannelPath\Office" -Force -Recurse
              }
           }

           $cabFilePath = "$env:TEMP/ofl.cab"
           if (!(Test-Path $cabFilePath)) {
                WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Copying $LocalPath\ofl.cab to $cabFilePath" -LogFilePath $LogFilePath
                Copy-Item -Path "$LocalPath\ofl.cab" -Destination $cabFilePath -Force
           }

           CreateMainCabFiles -LocalPath $LocalPath -ChannelShortName $ChannelShortName -LatestVersion $latestVersion

           $DeploymentFilePath = "$PSSCriptRoot\DeploymentFiles\*.*"
           if (Test-Path -Path $DeploymentFilePath) {
             WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Copying $DeploymentFilePath to $LocalPath" -LogFilePath $LogFilePath
             Copy-Item -Path $DeploymentFilePath -Destination "$LocalPath" -Force -Recurse
           } else {
             WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Deployment folder missing: $DeploymentFilePath" -LogFilePath $LogFilePath
             throw "Deployment folder missing: $DeploymentFilePath"
           }

           LoadCMPrereqs -SiteCode $SiteCode -CMPSModulePath $CMPSModulePath

           if (!($existingPackage)) {
              $package = CreateCMPackage -Name $packageName -Path $Path -Channel $Channel -UpdateOnlyChangedBits $UpdateOnlyChangedBits -CustomPackageShareName $CustomPackageShareName -LogFilePath $LogFilePath
              $packageCreated = $true
              Write-Host "`tPackage Created: $packageName"
              WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Package Created: $packageName" -LogFilePath $LogFilePath
           } else {
              if(!$packageCreated){
                Write-Host "`tPackage Already Exists: $packageName"
                WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Package Already Exists: $packageName" -LogFilePath $LogFilePath
              }
           }

           Write-Host

         }
       }
       } catch {
         throw;
       }
    }
    End
    {
        Set-ExecutionPolicy $currentExecutionPolicy -Scope Process -Force
        Set-Location $startLocation    
    }
}

function Update-CMOfficePackage {
<#
.SYNOPSIS
Automates the configuration of System Center Configuration Manager (CM) to update the Office Click-To-Run package

.DESCRIPTION


.PARAMETER Channels
The update channel. Current, Deferred, FirstReleaseDeferred, FirstReleaseCurrent

.PARAMETER OfficeSourceFilesPath
The location of the source files.

.PARAMETER MoveSourceFiles
This moves the files from the Source location to the location specified.

.PARAMETER SiteCode
The site code you would like to create the package on. If left blank it will default to the current site.

.PARAMETER CMPSModulePath
Allows the user to specify that full path to the ConfigurationManager.psd1 PowerShell Module. This is especially useful if CM is installed in a non standard path.

.PARAMETER UpdateDistributionPoints
Sets the distribution point to update to the latest files in the package share.


.EXAMPLE
Update-CMOfficePackage -Channels Deferred -Bitness v32 -OfficeSourceFilesPath D:\OfficeChannelFiles

.EXAMPLE
Update-CMOfficePackage -Channels Current -Bitness Both -OfficeSourceFilesPath D:\OfficeChannelFiles -UpdateDistributionPoints


#>   
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param
    (
        [Parameter()]
        [OfficeChannel[]] $Channels = @(0,1,2,3,4,5,6,7),

        [Parameter()]
	    [String]$OfficeSourceFilesPath = $NULL,

        [Parameter()]
	    [bool]$MoveSourceFiles = $false,

	    [Parameter()]
	    [String]$SiteCode = $null,

	    [Parameter()]
	    [String]$CMPSModulePath = $NULL,

        [Parameter()]
	    [bool]$UpdateDistributionPoints = $true,

        [Parameter()]
        [bool]$WaitForUpdateToFinish = $false,

        [Parameter()]
        [string]$LogFilePath
    )
    Begin
    {
        $currentExecutionPolicy = Get-ExecutionPolicy
	    Set-ExecutionPolicy Unrestricted -Scope Process -Force  
        $startLocation = Get-Location
    }
    Process {
       try {
       $currentFileName = Get-CurrentFileName
       Set-Alias -name LINENUM -value Get-CurrentLineNumber

       Check-AdminAccess

       $cabFilePath = "$OfficeSourceFilesPath\ofl.cab"
       if (Test-Path $cabFilePath) {
            Copy-Item -Path $cabFilePath -Destination "$PSScriptRoot\ofl.cab" -Force
       }

       $ChannelList = @("FirstReleaseCurrent", "Current", "FirstReleaseDeferred", "Deferred","MonthlyTargeted","Monthly","SemiAnnualTargeted","SemiAnnual")
       $ChannelXml = Get-ChannelXml -FolderPath $OfficeSourceFilesPath -OverWrite $false
       [bool]$packageNotification = $false

       foreach ($Channel in $ChannelList) {
         if ($Channels -contains $Channel) {
           switch($Channel){
               "FirstReleaseCurrent"{
                   $cabChannelName = ""
               }
               "Current"{
                   $cabChannelName = "Monthly"
               }
               "FirstReleaseDeferred"{
                   $cabChannelName = "Targeted"
               }
               "Deferred"{
                   $cabChannelName = "Broad"
               }
               "MonthlyTargeted"{
                   $cabChannelName = "Insiders"
               }
               "Monthly"{
                   $cabChannelName = "Monthly"
               }
               "SemiAnnualTargeted"{
                   $cabChannelName = "Targeted"
               }
               "SemiAnnual"{
                   $cabChannelName = "Broad"
               }
           }

           $selectChannel = $ChannelXml.UpdateFiles.baseURL | Where {$_.branch -eq $cabChannelName.ToString() }
           $latestVersion = Get-ChannelLatestVersion -ChannelUrl $selectChannel.URL -Channel $cabChannelName -FolderPath $OfficeFilesPath -OverWrite $false

           $ChannelShortName = ConvertChannelNameToShortName -ChannelName $Channel
           $existingPackage = CheckIfPackageExists
           if (!($existingPackage)) {
              WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "No Package Exists to Update. Please run the Create-CMOfficePackage function first to create the package." -LogFilePath $LogFilePath
              throw "No Package Exists to Update. Please run the Create-CMOfficePackage function first to create the package."
           }

           $packagePath = $existingPackage.PkgSourcePath
           if ($packagePath.StartsWith("\\")) {
               $shareName = $packagePath.Split("\")[3]
           }

           $existingShare = Get-Fileshare -Name $shareName
           if (!($existingShare)) {
              WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "No Package Exists to Update. Please run the Create-CMOfficePackage function first to create the package." -LogFilePath $LogFilePath
              throw "No Package Exists to Update. Please run the Create-CMOfficePackage function first to create the package."
           }

           $packageName = $existingPackage.Name
           $packageId = $existingPackage.PackageID

           if(!$packageNotification){
               Write-Host "`tUpdating Package: $packageName"
               WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Updating Package: $packageName" -LogFilePath $LogFilePath
               $packageNotification = $true
           }           

           $Path = $existingPackage.PkgSourcePath

           $packageName = "Office 365 ProPlus"
           $ChannelPath = "$Path\$Channel"
           $LocalPath = $existingShare.Path
           $LocalChannelPath = $existingShare.Path + "\SourceFiles"

           [System.IO.Directory]::CreateDirectory($LocalChannelPath) | Out-Null
                          
           if ($OfficeSourceFilesPath) {
                Write-Host "`t`tUpdating Source Files for $Channel..."
                WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Updating Source Files for $Channel..." -LogFilePath $LogFilePath

                $officeFileChannelPath = "$OfficeSourceFilesPath\$ChannelShortName"
                $officeFileTargetPath = "$LocalChannelPath\$ChannelShortName"

                if (!(Test-Path -Path $officeFileChannelPath)) {
                    WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Channel Folder Missing: $officeFileChannelPath - Ensure that you have downloaded the Channel you are trying to deploy" -LogFilePath $LogFilePath
                    throw "Channel Folder Missing: $officeFileChannelPath - Ensure that you have downloaded the Channel you are trying to deploy"
                }

                $tempofficeFileChannelPath = "$officeFileChannelPath\Office\Data"
                $tempLocalChannelPath = "$LocalChannelPath\$ChannelShortName\Office\Data"

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
                        #[System.IO.Directory]::CreateDirectory($officeFileTargetPath) | Out-Null
                        WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Moving $officeFileChannelPath to $LocalChannelPath" -LogFilePath $LogFilePath
                        Move-Item -Path $officeFileChannelPath -Destination $LocalChannelPath -Force
                    }else{
                        $subfiles = Get-ChildItem $tempofficeFileChannelPath
                        foreach($file in $subfiles){
                            [array]$tempLocalChannelPathFiles = (Get-ChildItem -Path $tempLocalChannelPath).Name
                            if($tempLocalChannelPathFiles -notcontains $file.Name){
                                WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Moving $tempofficeFileChannelPath\$file to $tempLocalChannelPath" -LogFilePath $LogFilePath
                                Move-Item -Path $tempofficeFileChannelPath\$file -Destination $tempLocalChannelPath -Force
                            }
                            else{
                                [array]$versionFiles = (Get-ChildItem -Path $tempLocalChannelPath\$latestVersion).Name
                                [array]$officeChannelVersionFiles = (Get-ChildItem -Path "$tempofficeFileChannelPath\$latestVersion").Name
                                foreach($officeChannelVersionFile in $officeChannelVersionFiles) {
                                    if($versionFiles -notcontains $officeChannelVersionFile){
                                        WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Moving $tempofficeFileChannelPath\$latestVersion\$officeChannelVersionFile to $tempLocalChannelPath\$latestVersion" -LogFilePath $LogFilePath
                                        Move-Item -Path $tempofficeFileChannelPath\$latestVersion\$officeChannelVersionFile -Destination $tempLocalChannelPath\$latestVersion -Force
                                    }
                                }
                            }           
                        }

                        Get-ChildItem -Path $officeFileChannelPath -Recurse | Remove-Item -Force -Recurse | Out-Null

                        [System.IO.Directory]::Delete($officeFileChannelPath) | Out-Null
                    }
                }else {
                    if(!(Test-Path -Path $officeFileTargetPath)) {
                        [System.IO.Directory]::CreateDirectory($officeFileTargetPath) | Out-Null
                        WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Copying $officeFileChannelPath to $LocalChannelPath" -LogFilePath $LogFilePath
                        Copy-Item -Path $officeFileChannelPath -Destination $LocalChannelPath -Recurse -Force
                    }
                    else{
                        $subfiles = Get-ChildItem $tempofficeFileChannelPath
                        foreach($file in $subfiles){
                            WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Copying $tempofficeFileChannelPath\$file to $tempLocalChannelPath" -LogFilePath $LogFilePath
                            Copy-Item -Path $tempofficeFileChannelPath\$file -Destination $tempLocalChannelPath -Recurse -Force 
                        }             
                    }
                }

                $cabFilePath = "$OfficeSourceFilesPath\ofl.cab"
                if (Test-Path $cabFilePath) {
                    WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Copying $cabFilePath to $LocalPath\ofl.cab" -LogFilePath $LogFilePath
                    Copy-Item -Path $cabFilePath -Destination "$LocalPath\ofl.cab" -Force
                }

           }

           $cabFilePath = "$env:TEMP/ofl.cab"
           if (!(Test-Path $cabFilePath)) {
                WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Copying $LocalPath\ofl.cab to $cabFilePath" -LogFilePath $LogFilePath
                Copy-Item -Path "$LocalPath\ofl.cab" -Destination $cabFilePath -Force
           }

           CreateMainCabFiles -LocalPath $LocalPath -ChannelShortName $ChannelShortName -LatestVersion $latestVersion

         }
       }

       $DeploymentFilePath = "$PSSCriptRoot\DeploymentFiles\*.*"
           if (Test-Path -Path $DeploymentFilePath) {
             Write-Host "`t`tUpdating Deployment Files..."
             WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Updating Deployment Files..." -LogFilePath $LogFilePath
             WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Copying $DeploymentFilePath to $LocalPath" -LogFilePath $LogFilePath
             Copy-Item -Path $DeploymentFilePath -Destination "$LocalPath" -Force -Recurse
           } else {
             WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Deployment folder missing: $DeploymentFilePath" -LogFilePath $LogFilePath
             throw "Deployment folder missing: $DeploymentFilePath"
           }

       LoadCMPrereqs -SiteCode $SiteCode -CMPSModulePath $CMPSModulePath

       if ($UpdateDistributionPoints) {
           Write-Host "`t`tUpdating Distribution Points..."
           WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Updating Distribution Points..." -LogFilePath $LogFilePath
           Update-CMDistributionPoint -PackageId $packageId
           if($WaitForUpdateToFinish){
               $distributionStatus = Get-CMOfficeDistributionStatus
               if(!$distributionStatus){
                   Write-Host ""
                   Write-Host "NOTE: In order to update the package you must run the function 'Distribute-CMOfficePackage'." -BackgroundColor Red
                   Write-Host "      You should wait until the content has finished distributing to the distribution points." -BackgroundColor Red
                   Write-Host "      Otherwise the deployments will fail. The clients will continue to fail until the " -BackgroundColor Red
                   Write-Host "      content distribution is complete." -BackgroundColor Red
               }

               Get-CMOfficeDistributionStatus -WaitForDistributionToFinish $true -LogFilePath $LogFilePath
           }
       }

       } catch {
         throw;
       }
    }
    End
    {
        Set-ExecutionPolicy $currentExecutionPolicy -Scope Process -Force
        Set-Location $startLocation    
    }
}

function Create-CMOfficeDeploymentProgram {
<#
.SYNOPSIS
Automates the configuration of System Center Configuration Manager (CM) to configure Office Click-To-Run Deployment

.DESCRIPTION
Creates a program that can be deployed to clients in a target collection to install Office 365 ProPlus.

.PARAMETER Channels
The update channel. Current, Deferred, FirstReleaseDeferred, FirstReleaseCurrent

.PARAMETER Bitness
Downloads the bitness of Office Click-to-Run "v32, v64, Both"

.PARAMETER DeploymentType
Chose how you would like to deploy Office. DeployWithScript, DeployWithConfigurationFile

.PARAMETER ScriptName
Name the script you would like to use "configuration.xml"

.PARAMETER SiteCode 
The site code you would like to create the package on. If left blank it will default to the current site

.PARAMETER CMPSModulePath
Allows the user to specify that full path to the ConfigurationManager.psd1 PowerShell Module. This is especially useful if CM is installed in a non standard path.

.PARAMETER ConfigurationXml
Sets the configuration file to be used for the instalation.

.PARAMETER CustomName
Replaces the default program name with a custom name. The custom name will also need to be provided when running the Deploy-CMOfficeProgram function.

.EXAMPLE
Create-CMOfficeDeploymentProgram -Channels Deferred -DeploymentType DeployWithScript

.EXAMPLE
Create-CMOfficeDeploymentProgram -Channels Current -DeploymentType DeployWithConfigurationFile -ScriptName engineering.xml

#>
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param
    (
        [Parameter()]
        [OfficeChannel[]] $Channels = @(0,1,2,3,4,5,6,7),

        [Parameter()]
	    [Bitness]$Bitness = "v32",

	    [Parameter()]
	    [CMDeploymentType]$DeploymentType = 0,

	    [Parameter()]
	    [String]$ScriptName = "CM-OfficeDeploymentScript.ps1",

	    [Parameter()]
	    [String]$SiteCode = $null,

	    [Parameter()]
	    [String]$CMPSModulePath = $NULL,

	    [Parameter()]
	    [String]$ConfigurationXml = ".\DeploymentFiles\DefaultConfiguration.xml",

	    [Parameter()]
	    [String]$CustomName = $NULL,

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
       try {
         $currentFileName = Get-CurrentFileName
         Set-Alias -name LINENUM -value Get-CurrentLineNumber

         Check-AdminAccess

         $platforms = @()

         if ($Bitness -eq "Both") {
            $platforms += "32"
            $platforms += "64"
         } else {
            $platforms += $Bitness.ToString().Replace("v", "")
         }

         if ($ConfigurationXml.StartsWith(".\")) {
           $ConfigurationXml = $ConfigurationXml -replace "^\.\\", "$PSScriptRoot\"
         }

         $tmpCustomName = $CustomName
         $CustomName = $CustomName -replace ' ', ''

         if ($CustomName) {
             if (!($CustomName.ToLower().Contains("deploy"))) {
                $CustomName = "Deploy " + $CustomName
                $tmpCustomName = "Deploy " + $tmpCustomName
             }
         }

         $CustomName = $CustomName -replace ' ', ''

         if ($CustomName.Length -gt 50) {
             WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "CustomName is too long.  Must be less then 50 Characters" -LogFilePath $LogFilePath
             throw "CustomName is too long.  Must be less then 50 Characters"
         }

         foreach ($channel in $Channels) {
             LoadCMPrereqs -SiteCode $SiteCode -CMPSModulePath $CMPSModulePath

             $LargeDrv = Get-LargestDrive
             $LocalPath = "$LargeDrv\OfficeDeployment"

             $channelShortName = ConvertChannelNameToShortName -ChannelName $channel

             $existingPackage = CheckIfPackageExists
             if (!($existingPackage)) {
                WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError  "You must run the Create-CMOfficePackage function before running this function" -LogFilePath $LogFilePath
                throw "You must run the Create-CMOfficePackage function before running this function"
             }

             foreach ($platform in $platforms) {
                 [string]$CommandLine = ""
                 [string]$ProgramName = ""

                 [string]$channelShortNameLabel = $channel
                 if ($channel -eq "FirstReleaseCurrent") {
                    $channelShortNameLabel = "FRCC"
                 }
                 if ($channel -eq "FirstReleaseDeferred") {
                    $channelShortNameLabel = "FRDC"
                 }

                 if ($DeploymentType -eq "DeployWithScript") {
                     $ProgramName = "Deploy $channelShortNameLabel Channel With Script - $platform-Bit"
                     if ($CustomName) {
                       $ProgramName = $tmpCustomName
                     }

                     $CommandLine = "%windir%\Sysnative\windowsPowershell\V1.0\powershell.exe -ExecutionPolicy Bypass -NoLogo -NonInteractive " + `
                                    "-NoProfile -WindowStyle Hidden -Command .\$ScriptName -Channel $channel -SourceFileFolder SourceFiles -Bitness $platform"

                 } elseif ($DeploymentType -eq "DeployWithConfigurationFile") {
                     if (!(Test-Path -Path $ConfigurationXml)) {
                        WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError  "Configuration file does not exist: $ConfigurationXml" -LogFilePath $LogFilePath
                        throw "Configuration file does not exist: $ConfigurationXml"
                     }

                     #[guid]::NewGuid().Guid.ToString()

                     $configId = "Config-$channel-$platform-Bit"
                     $configFileName = $configId + ".xml"

                     if ($CustomName) {
                         $configFileName = $configId + "-" + $CustomName + ".xml"
                     }

                     $configFilePath = "$LocalPath\$configFileName"

                     WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError  "Copying $ConfigurationXml to $configFilePath" -LogFilePath $LogFilePath
                     Copy-Item -Path $ConfigurationXml -Destination $configFilePath

                     $sourcePath = $NULL

                     $sourceFilePath = "$LocalPath\SourceFiles\$channelShortName\Office\Data"
                     if (Test-Path -Path $sourceFilePath) {
                        $sourcePath = ".\SourceFiles\$channelShortName"
                     } else {
                       $sourceFilePath = "$LocalPath\SourceFiles\$channel\Office\Data"
                       if (Test-Path -Path $sourceFilePath) {
                          $sourcePath = ".\SourceFiles\$channel"
                       }
                     }

                     UpdateConfigurationXml -Path $configFilePath -Channel $channel -Bitness $platform -SourcePath $sourcePath -LogFilePath $LogFilePath

                     $ProgramName = "Deploy $channelShortName with Config File - $platform-Bit"

                     if ($CustomName) {
                       $ProgramName = $tmpCustomName
                     }

                     $CommandLine = "%windir%\Sysnative\windowsPowershell\V1.0\powershell.exe -ExecutionPolicy Bypass -NoLogo -NonInteractive " + `
                                    "-NoProfile -WindowStyle Hidden -Command .\DeployConfigFile.ps1 -ConfigFileName $configFileName"
                 }

                 [string]$packageId = $null

                 $packageId = $existingPackage.PackageId
                 if ($packageId) {
                    $comment = $DeploymentType.ToString() + "-" + $channel + "-" + $platform

                    if ($CustomName) {
                       $comment += "-$CustomName"
                    }

                    WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError  "Creating the program: $ProgramName" -LogFilePath $LogFilePath
                    CreateCMProgram -Name $ProgramName -PackageID $packageId -RequiredPlatformNames $requiredPlatformNames -CommandLine $CommandLine -Comment $comment -LogFilePath $LogFilePath
                 }
             }
         }

       } catch {
         throw;
       }
    }
    End
    {
        Set-ExecutionPolicy $currentExecutionPolicy -Scope Process -Force
        Set-Location $startLocation    
    }
}

function Create-CMOfficeChannelChangeProgram {
 <#
.SYNOPSIS
Automates the configuration of System Center Configuration Manager (CM) to configure Office Click-To-Run channel change program

.DESCRIPTION
Creates an Office 365 ProPlus program that will change the channel of the client in a target collection.

.PARAMETER Channels
The update channel. Current, Deferred, FirstReleaseDeferred, FirstReleaseCurrent

.PARAMETER SiteCode
The 3 Letter Site ID.

.PARAMETER CMPSModulePath
Allows the user to specify that full path to the ConfigurationManager.psd1 PowerShell Module. This is especially useful if CM is installed in a non standard path.

.Example
Create-CMOfficeChannelChangeProgram -Sitecode S01 -Channels Current
#>
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param
    (
        [Parameter()]
        [OfficeChannel[]] $Channels = @(0,1,2,3,4,5,6,7),

	    [Parameter()]
	    [String]$SiteCode = $null,

	    [Parameter()]
	    [String]$CMPSModulePath = $NULL,

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
       try {
         $currentFileName = Get-CurrentFileName
         Set-Alias -name LINENUM -value Get-CurrentLineNumber

         Check-AdminAccess

         LoadCMPrereqs -SiteCode $SiteCode -CMPSModulePath $CMPSModulePath

         $existingPackage = CheckIfPackageExists
         if (!($existingPackage)) {
            WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError  "You must run the Create-CMOfficePackage function before running this function" -LogFilePath $LogFilePath
            throw "You must run the Create-CMOfficePackage function before running this function"
         }

         [string]$CommandLine = ""
         [string]$ProgramName = ""

         foreach ($channel in $Channels) {
             $ProgramName = "Change Channel to $channel"
             $CommandLine = "%windir%\Sysnative\windowsPowershell\V1.0\Powershell.exe -ExecutionPolicy Bypass -NoLogo -NonInteractive -NoProfile -WindowStyle Hidden -Command .\Change-OfficeChannel.ps1 -Channel $Channel"

             $SharePath = $existingPackage.PkgSourcePath

             $OSSourcePath = "$PSScriptRoot\DeploymentFiles\Change-OfficeChannel.ps1"
             $OCScriptPath = "$SharePath\Change-OfficeChannel.ps1"

             if (!(Test-Path $OSSourcePath)) {
                WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Required file missing: $OSSourcePath" -LogFilePath $LogFilePath
                throw "Required file missing: $OSSourcePath"
             } else {
                 if (!(Test-ItemPathUNC -Path $SharePath -FileName "Change-OfficeChannel.ps1")) {
                    Copy-ItemUNC -SourcePath $OSSourcePath -TargetPath $SharePath -FileName "Change-OfficeChannel.ps1"
                 }

                 [string]$packageId = $existingPackage.PackageId
                 if ($packageId) {
                    $comment = "ChangeChannel-$channel"

                    WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Creating the program: $ProgramName" -LogFilePath $LogFilePath
                    CreateCMProgram -Name $ProgramName -PackageID $packageId -CommandLine $CommandLine -RequiredPlatformNames $requiredPlatformNames -Comment $comment -LogFilePath $LogFilePath
                 }
             }
         }

       } catch {
         throw;
       }
    }
    End
    {
        Set-ExecutionPolicy $currentExecutionPolicy -Scope Process -Force
        Set-Location $startLocation    
    }
}

function Create-CMOfficeRollBackProgram {
 <#
.SYNOPSIS
Automates the configuration of System Center Configuration Manager (CM) to configure Office Click-To-Run rollback program

.DESCRIPTION
Creates an Office 365 ProPlus program that will look at the update source and install the previous version.

.PARAMETER $SiteCode
The 3 Letter Site ID.

.PARAMETER CMPSModulePath
Allows the user to specify that full path to the ConfigurationManager.psd1 PowerShell Module. This is especially useful if CM is installed in a non standard path.

.Example
Create-CMOfficeRollBackProgram -Sitecode S01
#>  
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param
    (
	    [Parameter()]
	    [String]$SiteCode = $null,

	    [Parameter()]
	    [String]$CMPSModulePath = $NULL,

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
       try {
         $currentFileName = Get-CurrentFileName
         Set-Alias -name LINENUM -value Get-CurrentLineNumber

         Check-AdminAccess

         LoadCMPrereqs -SiteCode $SiteCode -CMPSModulePath $CMPSModulePath

         $existingPackage = CheckIfPackageExists
         if (!($existingPackage)) {
            WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "You must run the Create-CMOfficePackage function before running this function" -LogFilePath $LogFilePath
            throw "You must run the Create-CMOfficePackage function before running this function"
         }

         [string]$CommandLine = ""
         [string]$ProgramName = ""

         $ProgramName = "Rollback"
         $CommandLine = "%windir%\Sysnative\windowsPowershell\V1.0\Powershell.exe -ExecutionPolicy Bypass -NoLogo -NonInteractive -NoProfile -WindowStyle Hidden -Command .\Change-OfficeChannel.ps1 -Rollback"

         $SharePath = $existingPackage.PkgSourcePath

         $OSSourcePath = "$PSScriptRoot\DeploymentFiles\Change-OfficeChannel.ps1"
         $OCScriptPath = "$SharePath\Change-OfficeChannel.ps1"

         if (!(Test-Path $OSSourcePath)) {
            WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Required file missing: $OSSourcePath" -LogFilePath $LogFilePath
            throw "Required file missing: $OSSourcePath"
         } else {
             if (!(Test-ItemPathUNC -Path $SharePath -FileName "Change-OfficeChannel.ps1")) {
                Copy-ItemUNC -SourcePath $OSSourcePath -TargetPath $SharePath -FileName "Change-OfficeChannel.ps1"
             }

             [string]$packageId = $existingPackage.PackageId
             if ($packageId) {
                $comment = "RollBack"

                WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Creating the program: $ProgramName" -LogFilePath $LogFilePath
                CreateCMProgram -Name $ProgramName -PackageID $packageId -CommandLine $CommandLine -RequiredPlatformNames $requiredPlatformNames -Comment $comment -LogFilePath $LogFilePath
             }
         }

       } catch {
         throw;
       }
    }
    End
    {
        Set-ExecutionPolicy $currentExecutionPolicy -Scope Process -Force
        Set-Location $startLocation    
    }
}

function Create-CMOfficeUpdateProgram {
<#
.SYNOPSIS
Automates the configuration of System Center Configuration Manager (CM) to configure Office Click-To-Run rollback program

.DESCRIPTION
Creates an Office 365 ProPlus update program.

.PARAMETER WaitForUpdateToFinish
The PowerShell service will continue to run until the update has finished.

.PARAMETER EnableUpdateAnywhere
Attempts to update Office 365 ProPlus using the existing update source. If the update source is not available (mobile users) then
the script will failover to the CDN as an update source.

.PARAMETER ForceAppShutdown
If set to $true Office apps will close automatically.

.PARAMETER UpdatePromptUser
If set to $true the user will be prompted to update.

.PARAMETER DisplayLevel
If ste to $true the update will be visible.

.PARAMETER UpdateToVersion
The version to update to.

.PARAMETER SiteCode
The 3 letter site code.

.PARAMETER CMPSModulePath
Allows the user to specify that full path to the ConfigurationManager.psd1 PowerShell Module. This is especially useful if CM is installed in a non standard path.

.PARAMETER UseScriptLocationAsUpdateSource
The location where the script is ran will be the location of the update source files.

.EXAMPLE
Create-CMOfficeUpdateProgram

.EXAMPLE
Create-CMOfficeUpdateProgram -ForceAppShutdown $true -EnableUpdateAnywhere $false -LogPath "$env:PUBLIC\UpdateOffice.log" -UpdateToVersion 16.0.6001.1078

#>
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param
    (
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
        [string] $UpdateToVersion = $NULL,

        [Parameter()]
        [string] $LogPath = $NULL,

        [Parameter()]
        [string] $LogName = $NULL,
        
        [Parameter()]
        [bool] $ValidateUpdateSourceFiles = $true,

	    [Parameter()]
	    [String]$SiteCode = $null,

	    [Parameter()]
	    [String]$CMPSModulePath = $NULL,

	    [Parameter()]
	    [bool]$UseScriptLocationAsUpdateSource = $true,

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
       try {
         $currentFileName = Get-CurrentFileName
         Set-Alias -name LINENUM -value Get-CurrentLineNumber

         Check-AdminAccess

         LoadCMPrereqs -SiteCode $SiteCode -CMPSModulePath $CMPSModulePath

         $existingPackage = CheckIfPackageExists
         if (!($existingPackage)) {
            WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "You must run the Create-CMOfficePackage function before running this function" -LogFilePath $LogFilePath
            throw "You must run the Create-CMOfficePackage function before running this function"
         }

         [string]$ProgramName = "Update Office 365 With ConfigMgr"
         [string]$CommandLine = "%windir%\Sysnative\windowsPowershell\V1.0\Powershell.exe -ExecutionPolicy Bypass -NoLogo -NonInteractive -NoProfile -WindowStyle Hidden -Command .\Update-Office365Anywhere.ps1"

         $CommandLine += " -WaitForUpdateToFinish " + (Convert-Bool -value $WaitForUpdateToFinish) + ` 
                         " -EnableUpdateAnywhere " + (Convert-Bool -value $EnableUpdateAnywhere) + ` 
                         " -ForceAppShutdown " + (Convert-Bool -value $ForceAppShutdown) + ` 
                         " -UpdatePromptUser " + (Convert-Bool -value $UpdatePromptUser) + ` 
                         " -DisplayLevel " + (Convert-Bool -value $DisplayLevel) + ` 
                         " -UseScriptLocationAsUpdateSource " + (Convert-Bool -value $UseScriptLocationAsUpdateSource)

         if ($UpdateToVersion) {
             $CommandLine += " -UpdateToVersion " + $UpdateToVersion
         }

         $SharePath = $existingPackage.PkgSourcePath

         $OSSourcePath = "$PSScriptRoot\DeploymentFiles\Update-Office365Anywhere.ps1"
         $OCScriptPath = "$SharePath\Update-Office365Anywhere.ps1"

         if (!(Test-Path $OSSourcePath)) {
            WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Required file missing: $OSSourcePath" -LogFilePath $LogFilePath
            throw "Required file missing: $OSSourcePath"
         } else {
             if (!(Test-ItemPathUNC -Path $SharePath -FileName "Update-Office365Anywhere.ps1")) {
                Copy-ItemUNC -SourcePath $OSSourcePath -TargetPath $SharePath -FileName "Update-Office365Anywhere.ps1"
             }

             [string]$packageId = $existingPackage.PackageId
             if ($packageId) {
                $comment = "UpdateWithConfigMgr"

                WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Creating the program: $ProgramName" -LogFilePath $LogFilePath
                CreateCMProgram -Name $ProgramName -PackageID $packageId -CommandLine $CommandLine -RequiredPlatformNames $requiredPlatformNames -Comment $comment -LogFilePath $LogFilePath
             }
         }

       } catch {
         throw;
       }
    }
    End
    {
        Set-ExecutionPolicy $currentExecutionPolicy -Scope Process -Force
        Set-Location $startLocation    
    }
}

function Create-CMOfficeUpdateAsTaskProgram {
<#
.SYNOPSIS
Creates an Office 365 ProPlus program.

.DESCRIPTION
Creates an Office 365 ProPlus program that will create a scheduled task on clients in the target collection.

.PARAMETER WaitForUpdateToFinish
The PowerShell service will continue to run until the update has finished.

.PARAMETER EnableUpdateAnywhere
Attempts to update Office 365 ProPlus using the existing update source. If the update source is not available (mobile users) then
the script will failover to the CDN as an update source.

.PARAMETER ForceAppShutdown
If set to $true Office apps will close automatically.

.PARAMETER UpdatePromptUser
If set to $true the user will be prompted to update.

.PARAMETER DisplayLevel
If ste to $true the update will be visible.

.PARAMETER UpdateToVersion
The version to update to.

.PARAMETER UseRandomStartTime
A random start time for the scheduled task.

.PARAMETER RandomTimeEnd 
A random end time for the scheduled task.

.PARAMETER StartTime
The actual start time for the scheduled task.

.PARAMETER SiteCode
The 3 letter site code.

.PARAMETER CMPSModulePath 
Allows the user to specify that full path to the ConfigurationManager.psd1 PowerShell Module. This is especially useful if CM is 
installed in a non standard path.

.PARAMETER UseScriptLocationAsUpdateSource
The location where the script is ran will be the location of the update source files.

.EXAMPLE
Create-CMOfficeUpdateAsTaskProgram -UpdateToVersion 16.0.6001.1078
Creates an Office 365 ProPlus program called 'Update Office 365 With Scheduled Task' that will update the client to version 16.0.6001.1078.

.EXAMPLE
Create-CMOfficeUpdateAsTaskProgram -WaitForUpdateToFinish $true -EnableUpdateAnywhere $true -ForceAppShutdown $false -UpdatePromptUser $true -DisplayLevel $true -StartTime 12:00
Creates an Office 365 ProPlus program called 'Update Office 365 With Scheduled Task'. The program will run on clients in the target collection every Tuesday. The client will
be prompted before updating and will display the progress. 

#>
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param
    (
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
        [string] $UpdateToVersion = $NULL,

        [Parameter()]
        [bool] $UseRandomStartTime = $true,

        [Parameter()]
        [string] $RandomTimeStart = "08:00",

        [Parameter()]
        [string] $RandomTimeEnd = "17:00",

        [Parameter()]
        [string] $StartTime,

        [Parameter()]
        [string] $LogPath = $NULL,

        [Parameter()]
        [string] $LogName = $NULL,
        
        [Parameter()]
        [bool] $ValidateUpdateSourceFiles = $true,

	    [Parameter()]
	    [String]$SiteCode = $null,

	    [Parameter()]
	    [String]$CMPSModulePath = $NULL,

	    [Parameter()]
	    [bool]$UseScriptLocationAsUpdateSource = $true,

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
       try {
         $currentFileName = Get-CurrentFileName
         Set-Alias -name LINENUM -value Get-CurrentLineNumber

         Check-AdminAccess

         LoadCMPrereqs -SiteCode $SiteCode -CMPSModulePath $CMPSModulePath

         $existingPackage = CheckIfPackageExists
         if (!($existingPackage)) {
            WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "You must run the Create-CMOfficePackage function before running this function" -LogFilePath $LogFilePath
            throw "You must run the Create-CMOfficePackage function before running this function"
         }

         [string]$CommandLine = "%windir%\Sysnative\windowsPowershell\V1.0\Powershell.exe -ExecutionPolicy Bypass -NoLogo -NonInteractive -NoProfile -WindowStyle Hidden -Command .\Create-Office365AnywhereTask.ps1"
         [string]$ProgramName = "Update Office 365 With Scheduled Task"

         $CommandLine += " -WaitForUpdateToFinish " + (Convert-Bool -value $WaitForUpdateToFinish) + ` 
                         " -EnableUpdateAnywhere " + (Convert-Bool -value $EnableUpdateAnywhere) + ` 
                         " -ForceAppShutdown " + (Convert-Bool -value $ForceAppShutdown) + ` 
                         " -UpdatePromptUser " + (Convert-Bool -value $UpdatePromptUser) + ` 
                         " -DisplayLevel " + (Convert-Bool -value $DisplayLevel) + ` 
                         " -UseScriptLocationAsUpdateSource " + (Convert-Bool -value $UseScriptLocationAsUpdateSource)
                         
         if($UseRandomStartTime){
             $CommandLine += " -UseRandomStartTime " + (Convert-Bool -value $UseRandomStartTime) + `
                             " -RandomTimeStart " + $RandomTimeStart + `
                             " -RandomTimeEnd " + $RandomTimeEnd
         }

         if($StartTime){
             $CommandLine += " -StartTime " + $StartTime
         }

         if ($UpdateToVersion) {
             $CommandLine += " -UpdateToVersion " + $UpdateToVersion
         }

         $SharePath = $existingPackage.PkgSourcePath

         $OSSourcePath = "$PSScriptRoot\DeploymentFiles\Update-Office365Anywhere.ps1"
         $OCScriptPath = "$SharePath\Update-Office365Anywhere.ps1"

         $OSSourcePathTask = "$PSScriptRoot\DeploymentFiles\Create-Office365AnywhereTask.ps1"
         $OCScriptPathTask = "$SharePath\Create-Office365AnywhereTask.ps1"

         if (!(Test-Path $OSSourcePath)) {
            WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Required file missing: $OSSourcePath" -LogFilePath $LogFilePath
            throw "Required file missing: $OSSourcePath"
         } else {
             if (!(Test-ItemPathUNC -Path $SharePath -FileName "Update-Office365Anywhere.ps1")) {
                Copy-ItemUNC -SourcePath $OSSourcePath -TargetPath $SharePath -FileName "Update-Office365Anywhere.ps1"
             }

             if ($UseScheduledTask) {
               if (!(Test-ItemPathUNC -Path $SharePath -FileName "Create-Office365AnywhereTask.ps1")) {
                  Copy-ItemUNC -SourcePath $OSSourcePathTask  -TargetPath $SharePath -FileName "Create-Office365AnywhereTask.ps1"
               }
             }

             [string]$packageId = $existingPackage.PackageId
             if ($packageId) {
                $comment = "UpdateWithTask"

                WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Creating the program: $ProgramName" -LogFilePath $LogFilePath
                CreateCMProgram -Name $ProgramName -PackageID $packageId -CommandLine $CommandLine -RequiredPlatformNames $requiredPlatformNames -Comment $comment -LogFilePath $LogFilePath
             }
         }

       } catch {
         throw;
       }
    }
    End
    {
        Set-ExecutionPolicy $currentExecutionPolicy -Scope Process -Force
        Set-Location $startLocation    
    }
}

function Create-CMOfficeLanguageProgram{   
<#
.SYNOPSIS
Automates the configuration of System Center Configuration Manager (CM) to configure Office 365 ProPlus program deployments.

.DESCRIPTION
Creates a program in System Center Configuration Manager (CM) to deploy and install a specified language or multiple languages.

.PARAMETER Channel
The target update channel; Current, Deferred, FirstReleaseDeferred, or FirstReleaseCurrent.

.PARAMETER Languages
All office languages are supported in the ll-cc format "en-us"

.PARAMETER Bitness
The architecture of Office specified as v32 or v64.

.PARAMETER MainOfficeLanguage
The Shell UI language of Office.

.PARAMETER Version
The version of Office containing the language files.

.PARAMETER ConfigurationXml
Sets the configuration file to be used for the instalation.

.EXAMPLE
Create-CMOfficeLanguageProgram -Channel Deferred -Languages de-de,fr-fr -Bitness v32
A language pack configuration file will be created to add the German and French language packs. A program will be created
for the Office 365 ProPlus package that can be deployed to install additional languages on a client.

.EXAMPLE
Create-CMOfficeLanguageProgram -Channel Deferred -Languages de-de,fr-fr -Bitness v32 -MainOfficeLanguage ja-jp
A language pack configuration file will be created to add the German and French language packs. The Japanese language will
replace English as the Shell UI language. A program will be created for the Office 365 ProPlus package that can be deployed 
to install additional languages on a client

#>   
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param
    (
        [Parameter()]
        [OfficeChannel]$Channel = "Broad",

        [Parameter()]
        [ValidateSet("en-us","MatchOS","ar-sa","bg-bg","zh-cn","zh-tw","hr-hr","cs-cz","da-dk","nl-nl","et-ee","fi-fi","fr-fr","de-de","el-gr","he-il","hi-in","hu-hu","id-id","it-it",
                    "ja-jp","kk-kz","ko-kr","lv-lv","lt-lt","ms-my","nb-no","pl-pl","pt-br","pt-pt","ro-ro","ru-ru","sr-latn-rs","sk-sk","sl-si","es-es","sv-se","th-th",
                    "tr-tr","uk-ua","vi-vn")]
        [string[]]$Languages = ("en-us"),

        [Parameter()]
        [ValidateSet("af-za","sq-al","am-et","hy-am","as-in","az-latn-az","eu-es","be-by","bn-bd","bn-in","bs-latn-ba","ca-es","prs-af","fil-ph","gl-es","ka-ge","gu-in","is-is","ga-ie","kn-in",
                "km-kh","sw-ke","kok-in","ky-kg","lb-lu","mk-mk","ml-in","mt-mt","mi-nz","mr-in","mn-mn","ne-np","nn-no","or-in","fa-ir","pa-in","quz-pe","gd-gb","sr-cyrl-rs","sr-cyrl-ba",
                "sd-arab-pk","si-lk","ta-in","tt-ru","te-in","tk-tm","ur-pk","ug-cn","uz-latn-uz","ca-es-valencia","cy-gb")]
        [string[]] $PartialLanguages,

        [Parameter()]
        [ValidateSet("ha-latn-ng","ig-ng","xh-za","zu-za","rw-rw","ps-af","rm-ch","nso-za","tn-za","wo-sn","yo-ng")]
        [string[]] $ProofingLanguages,

        [Parameter()]
        [Bitness]$Bitness = "v32",

        [Parameter()]
        [string]$MainOfficeLanguage = "en-us",

        [Parameter()]
        [string]$Version = $NULL,

        [Parameter()]
	    [String]$ConfigurationXml = ".\DeploymentFiles\LanguageConfiguration.xml",

        [Parameter()]
        [string]$LogFilePath        
    )
    
    Begin {
        #create array for all languages including core, partial, and proofing
        $allLanguages = @();
        $Languages | 
        %{
          $allLanguages += $_
        }
	
        $PartialLanguages | 
        %{
          $allLanguages += $_
        }
	
        $ProofingLanguages | 
        %{
          $allLanguages += $_
        }
        $currentExecutionPolicy = Get-ExecutionPolicy
	    Set-ExecutionPolicy Unrestricted -Scope Process -Force  
        $startLocation = Get-Location
    }

    Process {
        try{
        $currentFileName = Get-CurrentFileName
        Set-Alias -name LINENUM -value Get-CurrentLineNumber

        Check-AdminAccess

        $LargeDrv = Get-LargestDrive
        $LocalPath = "$LargeDrv\OfficeDeployment"

        $ChannelList = @("FirstReleaseCurrent", "Current", "FirstReleaseDeferred", "Deferred","MonthlyTargeted","Monthly","SemiAnnualTargeted","SemiAnnual")
        $ChannelXml = Get-ChannelXml -FolderPath $LocalPath -OverWrite $true

        foreach ($ChannelName in $ChannelList) {
            if ($Channel -eq $ChannelName) {
                $selectChannel = $ChannelXml.UpdateFiles.baseURL | ? {$_.branch -eq $ChannelName.ToString() }
                $latestVersion = Get-ChannelLatestVersion -ChannelUrl $selectChannel.URL -Channel $ChannelName
                $ChannelShortName = ConvertChannelNameToShortName -ChannelName $ChannelName

                if ($Version) {
                    $latestVersion = $Version
                }

                $Bit = $Bitness.ToString().Replace("v", "")

                LoadCMPrereqs -SiteCode $SiteCode -CMPSModulePath $CMPSModulePath

                $siteDrive = Get-Location

                $existingPackage = CheckIfPackageExists
                if (!($existingPackage)) {
                    WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "You must run the Create-CMOfficePackage function before running this function" -LogFilePath $LogFilePath
                    throw "You must run the Create-CMOfficePackage function before running this function"
                }
        
                [string]$CommandLine = ""
                [string]$ProgramName = ""

                [string]$channelShortNameLabel = $ChannelName
                if ($Channel -eq "FirstReleaseCurrent") {
                    $channelShortNameLabel = "FRCC"
                }
                if ($Channel -eq "FirstReleaseDeferred") {
                    $channelShortNameLabel = "FRDC"
                }
                if ($Channel -eq "Current") {
                    $channelShortNameLabel = "CC"
                }
                if ($Channel -eq "Deferred") {
                    $channelShortNameLabel = "DC"
                }
                             
                $SharePath = $existingPackage.PkgSourcePath

                $OSSourcePath = "$PSScriptRoot\DeploymentFiles\DeployConfigFile.ps1"
                $OCScriptPath = "$SharePath\DeployConfigFile.ps1"

                $sourcePath = $NULL
                $sourceFilePath = "$LocalPath\SourceFiles\$channelShortName\Office\Data"
                if (Test-Path -Path $sourceFilePath) {
                    $sourcePath = ".\SourceFiles\$channelShortName"
                } else {
                    $sourceFilePath = "$LocalPath\SourceFiles\$Channel\Office\Data"
                    if (Test-Path -Path $sourceFilePath) {
                        $sourcePath = ".\SourceFiles\$Channel"
                    }
                }
                
                if($allLanguages.Count -gt "1"){
                    Set-Location $siteDrive

                    $languagePrograms = Get-CMProgram | ? {$_.ProgramName -like "DeployLanguagePack*" -and $_.ProgramName -like "*Multi*"}
                    if($languagePrograms){
                        $languageProgramNumList = @()
                        if($languagePrograms.Count -gt "0"){
                            foreach($Program in $languagePrograms){
                                $languageProgramName = $Program.ProgramName
                                $languageProgramNumList += $languageProgramName.Split("-")[4]
                            }

                            $languageProgramNumList = $languageProgramNumList | Sort-Object -Descending
                            $sortedProgramNum = $languageProgramNumList[0]
                            [int]$oldLanguageProgramNum = [convert]::ToInt32($sortedProgramNum, 10)                  
                            $newLanguageProgramNum = $oldLanguageProgramNum + 1

                            $ProgramName = "DeployLanguagePack-$Channel-" + $Bit + "bit-Multi-$newLanguageProgramNum"



                        }
                    } else {
                        $ProgramName = "DeployLanguagePack-$Channel-" + $Bit + "bit-Multi-1"    
                    }
                } else {
                    $ProgramName = "DeployLanguagePack-$Channel-" + "$Bit" + "bit-$allLanguages"
                }

                $configFileName = $ProgramName + ".xml"

                if ($CustomName) {
                    $configFileName = $ProgramName + "-" + $CustomName + ".xml"
                }

                $configFilePath = "$LocalPath\$configFileName"

                Set-Location $startLocation

                WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Copying $ConfigurationXml to $configFilePath" -LogFilePath $LogFilePath
                Copy-Item -Path $ConfigurationXml -Destination $configFilePath

                $CommandLine = "%windir%\Sysnative\windowsPowershell\V1.0\powershell.exe -ExecutionPolicy Bypass -NoLogo -NonInteractive " + `
                               "-NoProfile -WindowStyle Hidden -Command .\DeployConfigFile.ps1 -ConfigFileName $configFileName"
               
                if($MainOfficeLanguage.ToLower() -ne "en-us") {
                    [System.XML.XMLDocument]$ConfigFile = New-Object System.XML.XMLDocument
                    $ConfigFile.Load($configFilePath) | Out-Null
                    [System.XML.XMLElement]$nodes = $ConfigFile.Configuration.Add.Product.Language
                    foreach($node in $nodes) {
                        $node.SetAttribute("ID", "$MainOfficeLanguage")
                    }
                    $ConfigFile.Save($configFilePath)
                }
                
                foreach ($language in $allLanguages){
                    if(!(Get-ChildItem -Path $SharePath\SourceFiles\$channelShortName\Office\Data\$latestVersion | Where-Object {$_ -like "*$language*"})){
                        Remove-Item -Path $configFilePath
                        WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "The language pack $language was not found. To download the language run Download-CMOfficeChannelFiles using the -Languages parameter" -LogFilePath $LogFilePath
                        throw "The language pack $language was not found. To download the language run Download-CMOfficeChannelFiles using the -Languages parameter"
                    }
                    else{
                        UpdateConfigurationXml -Path $configFilePath -Channel $ChannelName -Bitness $Bit -SourcePath $sourcePath -Language $language -LogFilePath $LogFilePath
                    }
                }
        
                [string]$packageId = $null

                Set-Location $siteDrive

                $packageId = $existingPackage.PackageId
                if ($packageId) {
                    $comment = $NULL
                    foreach($language in $allLanguages){
                        if($comment -eq $NULL){
                            $comment += $language
                        } else {
                            $comment += ",$language"
                        }
                    }
                    #$comment = "DeployLanguagePack-$Channel-$Bit-$Languages"

                    if ($CustomName) {
                        $comment += "-$CustomName"
                    }

                    WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Creating the program: $ProgramName" -LogFilePath $LogFilePath
                    CreateCMProgram -Name $ProgramName -PackageID $packageId -RequiredPlatformNames $requiredPlatformNames -CommandLine $CommandLine -Comment $comment -LogFilePath $LogFilePath
                }
            }
        }

        } catch{
            throw;
        }
    }

    End {
        Set-ExecutionPolicy $currentExecutionPolicy -Scope Process -Force
        Set-Location $startLocation 
    }
}

function Distribute-CMOfficePackage {
<#
.SYNOPSIS
Automates the configuration of System Center Configuration Manager (CM) to configure Office Click-To-Run Updates

.DESCRIPTION
Distributes the Office 365 ProPlus package to the specified Distribution Point or Distribution Point Group.

.PARAMETER Channels
The update channel. Current, Deferred, FirstReleaseDeferred, FirstReleaseCurrent

.PARAMETER DistributionPoint
The distribution point name.

.PARAMETER DistributionPointGroupName
The distribution point group name.

.PARAMETER SiteCode
The 3 Letter Site ID.

.PARAMETER CMPSModulePath
Allows the user to specify that full path to the ConfigurationManager.psd1 PowerShell Module. This is especially useful if CM is installed in a non standard path.

.PARAMETER distributionPoint
Sets which distribution points will be used, and distributes the package.

.Example
Distribute-CMOfficePackage -DistirbutionPoint cm.contoso.com
Distributes the package 'Office 365 ProPlus' to the distribution point cm.contoso.com

#>
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param
    (
        [Parameter()]
        [OfficeChannel[]] $Channels = @(0,1,2,3,4,5,6,7),

	    [Parameter()]
	    [string]$DistributionPoint,

	    [Parameter()]
	    [string]$DistributionPointGroupName,

	    [Parameter()]
	    [String]$SiteCode = $null,

	    [Parameter()]
	    [String]$CMPSModulePath = $NULL,

        [Parameter()]
        [bool]$WaitForDistributionToFinish = $false,

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
       try {
        $currentFileName = Get-CurrentFileName
        Set-Alias -name LINENUM -value Get-CurrentLineNumber

        Check-AdminAccess

        $package = CheckIfPackageExists

        if (!($package)) {
            WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "You must run the Create-CMOfficePackage function before running this function" -LogFilePath $LogFilePath
            throw "You must run the Create-CMOfficePackage function before running this function"
        }

        LoadCMPrereqs -SiteCode $SiteCode -CMPSModulePath $CMPSModulePath

        if ($package) {
            [string]$packageName = $package.Name

            if ($DistributionPointGroupName) {
                Write-Host "Starting Content Distribution for package: $packageName"
                WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Starting Content Distribution for package: $packageName" -LogFilePath $LogFilePath
	            Start-CMContentDistribution -PackageName $packageName -DistributionPointGroupName $DistributionPointGroupName
            }

            if ($DistributionPoint) {
                Write-Host "Starting Content Distribution for package: $packageName"
                WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Starting Content Distribution for package: $packageName" -LogFilePath $LogFilePath
                Start-CMContentDistribution -PackageName $packageName -DistributionPointName $DistributionPoint
            }
        }

        if($WaitForDistributionToFinish){
            Get-CMOfficeDistributionStatus -WaitForDistributionToFinish $true -LogFilePath $LogFilePath
        }
        
        Write-Host 
        Write-Host "NOTE: In order to deploy the package you must run the function 'Deploy-CMOfficeChannelPackage'." -BackgroundColor Red
        Write-Host "      You should wait until the content has finished distributing to the distribution points." -BackgroundColor Red
        Write-Host "      otherwise the deployments will fail. The clients will continue to fail until the " -BackgroundColor Red
        Write-Host "      content distribution is complete." -BackgroundColor Red

       } catch {
         throw;
       }
    }
    End
    {
        Set-ExecutionPolicy $currentExecutionPolicy -Scope Process -Force
        Set-Location $startLocation    
    }
}

function Deploy-CMOfficeProgram {
<#
.SYNOPSIS
Automates the configuration of System Center Configuration Manager (CM) to configure Office 365 ProPlus program deployments.

.DESCRIPTION
Creates a deployment for the Office 365 ProPlus program created from the functions Create-CMOfficeDeploymentProgram, Create-CMOfficeChannelChangeProgram
Create-CMOfficeRollBackProgram, Create-CMOfficeUpdateProgram, and Create-CMOfficeUpdateAsTaskProgram.

.PARAMETER Collection
Required. The target ConfigMgr Collection ID.

.PARAMETER Channel
The target update channel; Current, Deferred, FirstReleaseDeferred, or FirstReleaseCurrent.

.PARAMETER ProgramType
Required. The type of program that will be deployed.
DeployWithScript
    A script will be used to configure the Office 365 ProPlus installation.
DeployWithConfigurationFile
    A configuration xml file will be used to install Office.
ChangeChannel
    The program is used to change the current channel of the client.
RollBack
    The program is used to rollback to a previous version.
UpdateWithConfigMg
    Configuration Manager will be used to start the update.
UpdateWithTask
    A task will be configured to start the update. 

.PARAMETER Bitness
The architecture of Office specified as v32 or v64.

.PARAMETER SiteCode
The 3 Letter Site ID.

.PARAMETER DeploymentPurpose
Choose between Default, Required, or Available. If the DeploymentPurpose is set to Required the program will be required to be installed on the client. If DeploymentPurpose
is Available the program will be become available on the client in the Software Center. If the DeploymentPurpose is not specified or set to Default the deployment will be
required to be installed on the clients of the collection.

.PARAMETER CMPSModulePath
Allows the user to specify that full path to the ConfigurationManager.psd1 PowerShell Module. This is especially useful if CM is installed in a non standard path.

.PARAMETER CustomName
Replaces the default program name with a custom name if it was provided while running Create-CMOfficeDeploymentProgram

.EXAMPLE
Deploy-CMOfficeProgram -Collection "Office Update" -ProgramType DeployWithScript
Creates an Office 365 ProPlus deployment for the program 'Office 365 ProPlus (Deploy Deferred Channel With Script - 32-Bit) that will be required to download on clients in the
target collection 'Office Update'.

.EXAMPLE
Deploy-CMOfficeProgram -Collection "Office Update" -ProgramType DeployWithScript -Channel Current -DeploymentPurpose Available
Creates an Office 365 ProPlus deployment for the program 'Office 365 ProPlus (Deploy Current Channel With Script - 32-Bit) that will be available in the software 
center on clients in the target collection 'Office Update'.

.EXAMPLE
Deploy-CMOfficeProgram -Collection "Office Update" -ProgramType DeployWithConfigurationFile
Creates an Office 365 ProPlus deployment for the program 'Office 365 ProPlus (Deploy Deffered Channel with Config File - 32-Bit). The deployment will be a 
required installation on clients in the target collection.

.EXAMPLE
Deploy-CMOfficeProgram -Collection "Office Update" -ProgramType DeployWithConfigurationFile -Channel FirstReleaseDeferred -DeploymentPurpose Available
Creates an Office 365 ProPlus deployment for the program 'Deploy FRDC Channel with Config File' that will be available in the software center on clients in the target 
collection 'Office Update'. The deployment will install Office 365 ProPlus FirstReleaseDeferred using an xml configuration file.

.EXAMPLE
Deploy-CMOfficeProgram -Collection "Office Update" -ProgramType ChangeChannel -Channel Current
Creates a deployment for the program 'Office 365 ProPlus (Change Channel to Current) that will be available in the software center on 
clients in the target collection 'Office Update'. 

.EXAMPLE
Deploy-CMOfficeProgram -Collection "Office Update" -ProgramType ChangeChannel -Channel FirstReleaseDeferred
Creates a deployment for the program 'Office 365 ProPlus (Change Channel to FirstReleaseDeferred) that will be available in the software center for 
clients in the target collection 'Office Update'.

.EXAMPLE
Deploy-CMOfficeProgram -Collection "Office Update" -ProgramType RollBack
Creates a deployment for the program 'Office 365 ProPlus (Rollback)' that will be required to download for clients in the
target collection 'Office Update'.

.EXAMPLE
Deploy-CMOfficeProgram -Collection "Office Update" -ProgramType RollBack -DeploymentPurpose Available
Creates a deployment for the program 'Office 365 ProPlus (Rollback)' that will be available in the software center for clients in the
target collection 'Office Update'.

.Example
Deploy-CMOfficeProgram -Collection "Office Update" -ProgramType UpdateWithConfigMgr -DeploymentPurpose Available -Channel Deferred
Deploys the Package created by the Setup-CMOfficeProPlusPackage function to Collection ID "Office Update".

.EXAMPLE
Deploy-CMOfficeProgram -Collection "Office Update" -ProgramType UpdateWithConfigMgr -DeploymentPurpose Available -Channel Deferred -Bitness v32
Deploys the Package created by the Setup-CMOfficeProPlusPackage function to Collection ID "Office Update" and will be referenced as 32 bit.

.EXAMPLE
Deploy-CMOfficeProgram -Collection "Office Update" -ProgramType UpdateWithTask
Creates a deployment for the program 'Office 365 ProPlus (Update Office 365 With Scheduled Task)' that will be required to download on clients in the
target collection 'Office Update'. 

.EXAMPLE
Deploy-CMOfficeProgram -Collection "Office Update" -ProgramType UpdateWithTask -Channel Deferred -DeploymentPurpose Available
Creates a deployment for the program 'Office 365 ProPlus (Update Office 365 With Scheduled Task)' that will be available in the software center for 
clients in the target collection 'Office Update'.

#>
    [CmdletBinding()]	
    Param
	(
		[Parameter(Mandatory=$true)]
		[String]$Collection = "",

        [Parameter(Mandatory=$true)]
        [CMOfficeProgramType] $ProgramType,

        [Parameter()]
        [OfficeChannel]$Channel = "Broad",

        [Parameter()]
	    [BitnessOptions]$Bitness = "v32",

        [Parameter()]
        [ValidateSet("en-us","MatchOS","ar-sa","bg-bg","zh-cn","zh-tw","hr-hr","cs-cz","da-dk","nl-nl","et-ee","fi-fi","fr-fr","de-de","el-gr","he-il","hi-in","hu-hu","id-id","it-it",
                    "ja-jp","kk-kz","ko-kr","lv-lv","lt-lt","ms-my","nb-no","pl-pl","pt-br","pt-pt","ro-ro","ru-ru","sr-latn-rs","sk-sk","sl-si","es-es","sv-se","th-th",
                    "tr-tr","uk-ua","vi-vn")]
        [string[]]$Languages = ("en-us"),

        [Parameter()]
        [ValidateSet("af-za","sq-al","am-et","hy-am","as-in","az-latn-az","eu-es","be-by","bn-bd","bn-in","bs-latn-ba","ca-es","prs-af","fil-ph","gl-es","ka-ge","gu-in","is-is","ga-ie","kn-in",
                "km-kh","sw-ke","kok-in","ky-kg","lb-lu","mk-mk","ml-in","mt-mt","mi-nz","mr-in","mn-mn","ne-np","nn-no","or-in","fa-ir","pa-in","quz-pe","gd-gb","sr-cyrl-rs","sr-cyrl-ba",
                "sd-arab-pk","si-lk","ta-in","tt-ru","te-in","tk-tm","ur-pk","ug-cn","uz-latn-uz","ca-es-valencia","cy-gb")]
        [string[]] $PartialLanguages,

        [Parameter()]
        [ValidateSet("ha-latn-ng","ig-ng","xh-za","zu-za","rw-rw","ps-af","rm-ch","nso-za","tn-za","wo-sn","yo-ng")]
        [string[]] $ProofingLanguages,
    
	    [Parameter()]
	    [String]$SiteCode = $NULL,

	    [Parameter()]
	    [String]$CMPSModulePath = $NULL,

    	[Parameter()]
	    [DeploymentPurpose]$DeploymentPurpose = "Default",

	    [Parameter()]
	    [String]$CustomName = $NULL,

        [Parameter()]
        [string]$LogFilePath
	) 
    Begin
    {
        $allLanguages = @();
        $allLanguages += , $Languages
        $allLanguages += , $PartialLanguages
        $allLanguages += , $ProofingLanguages
        $currentExecutionPolicy = Get-ExecutionPolicy
	    Set-ExecutionPolicy Unrestricted -Scope Process -Force  
        $startLocation = Get-Location
    }
    Process
    {
       try {
        $currentFileName = Get-CurrentFileName
        Set-Alias -name LINENUM -value Get-CurrentLineNumber

        Check-AdminAccess

        if ($CustomName) {
           if (!($CustomName.ToLower().Contains("deploy"))) {
              $CustomName = "Deploy " + $CustomName
           }
        }

        $CustomName = $CustomName -replace ' ', ''

        if ($CustomName.Length -gt 50) {
            WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "CustomName is too long.  Must be less then 50 Characters" -LogFilePath $LogFilePath
            throw "CustomName is too long.  Must be less then 50 Characters"
        }

        $strBitness = $Bitness.ToString() -Replace "v", ""

        $ChannelList = @("FirstReleaseCurrent", "Current", "FirstReleaseDeferred", "Deferred","MonthlyTargeted","Monthly","SemiAnnualTargeted","SemiAnnual")
        $ChannelXml = Get-ChannelXml

        foreach ($ChannelName in $ChannelList) {
            if ($Channel.ToString().ToLower() -eq $ChannelName.ToLower()) {
                switch($Channel){

                    "FirstReleaseCurrent"{
                        $cabChannelName = ""
                    }
                    "Current"{
                        $cabChannelName = "Monthly"
                    }
                    "FirstReleaseDeferred"{
                        $cabChannelName = "Targeted"
                    }
                    "Deferred"{
                        $cabChannelName = "Broad"
                    }
                    "MonthlyTargeted"{
                        $cabChannelName = "Insiders"
                    }
                    "Monthly"{
                        $cabChannelName = "Monthly"
                    }
                    "SemiAnnualTargeted"{
                        $cabChannelName = "Targeted"
                    }
                    "SemiAnnual"{
                        $cabChannelName = "Broad"
                    }
                }

                $selectChannel = $ChannelXml.UpdateFiles.baseURL | Where {$_.branch -eq $cabChannelName.ToString() }
                $latestVersion = Get-ChannelLatestVersion -ChannelUrl $selectChannel.URL -Channel $cabChannelName
                $ChannelShortName = ConvertChannelNameToShortName -ChannelName $ChannelName
                $package = CheckIfPackageExists

                if (!($package)) {
                    WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "You must run the Create-CMOfficePackage function before running this function" -LogFilePath $LogFilePath
                    throw "You must run the Create-CMOfficePackage function before running this function"
                }

                LoadCMPrereqs -SiteCode $SiteCode -CMPSModulePath $CMPSModulePath
                $SiteCode = GetLocalSiteCode -SiteCode $SiteCode

                $pType = ""

                Switch ($ProgramType) {
                    "DeployWithScript" { 
                         $pType = "DeployWithScript-$Channel-$strBitness"; 
                         if ($DeploymentPurpose -eq "Default") {
                             $DeploymentPurpose = "Required";
                         }
                    }
                    "DeployWithConfigurationFile" { 
                         $pType = "DeployWithConfigurationFile-$Channel-$strBitness"; 
                         if ($DeploymentPurpose -eq "Default") {
                           $DeploymentPurpose = "Required" 
                         } 
                    }
                    "ChangeChannel" { 
                         $pType = "ChangeChannel-$Channel"; 
                         if ($DeploymentPurpose -eq "Default") {
                             $DeploymentPurpose = "Available" 
                         }
                         $CustomName = $NULL
                    }
                    "RollBack" { 
                         $pType = "RollBack"; 
                         if ($DeploymentPurpose -eq "Default") {
                            $DeploymentPurpose = "Available" 
                         }
                         $CustomName = $NULL
                    }
                    "UpdateWithConfigMgr" { 
                         $pType = "UpdateWithConfigMgr"; 
                         if ($DeploymentPurpose -eq "Default") {
                            $DeploymentPurpose = "Required"  
                         }
                         $CustomName = $NULL
                    }
                    "UpdateWithTask" { 
                         $pType = "UpdateWithTask"; 
                         if ($DeploymentPurpose -eq "Default") {
                            $DeploymentPurpose = "Required"  
                         }
                         $CustomName = $NULL
                    }
                    "LanguagePack" {
                         $pType = "DeployLanguagePack-$Channel-$allLanguages";
                         if ($DeploymentPurpose -eq "Default") {
                            $DeploymentPurpose = "Available"  
                         }
                         $CustomName = $NULL
                    }
                }

                if ($DeploymentPurpose -eq "Default") {
                   $DeploymentPurpose = "Required"  
                }

                $tmpPType = $pType
                if ($CustomName) {
                   $tmpPType += "-$CustomName"
                }

                if($ProgramType -ne "LanguagePack") {
                    $Program = Get-CMProgram | Where {$_.Comment.ToLower() -eq $tmpPType.ToLower() }
                }
                else{
                    $languagePrograms = Get-CMProgram | ? {$_.ProgramName -like "DeployLanguage*"}
                    $tempLanguages = @()
                    foreach($lang in $allLanguages){
                        $tempLanguages += $lang
                    }
                    $tempLanguages = $tempLanguages | Sort-Object -Descending
                    [string]$sortedTempLanguages = $NULL
                    foreach($lang in $tempLanguages){
                        if(!$sortedTempLanguages){
                            $sortedTempLanguages += $lang
                        } else {
                            $sortedTempLanguages += ",$lang"
                        }
                    }
                    foreach($langProgram in $languagePrograms){
                        $commentLang = $langProgram.Comment.Split(",") | Sort-Object -Descending
                        [string]$sortedCommentLangs = $NULL
                        foreach($comLang in $commentLang){
                            if(!$sortedCommentLangs){
                                $sortedCommentLangs += $comLang
                            } else {
                                $sortedCommentLangs += ",$comLang"
                            }
                        }
                        if($langProgram.ProgramName -like "DeployLanguagePack-$Channel-$strBitness*" -and $sortedCommentLangs -eq $sortedTempLanguages){
                            $Program = $langProgram
                        }
                    }
                }

                $programName = $Program.ProgramName

                $packageName = "Office 365 ProPlus"
                if ($package) {
                   if ($Program) {

                        [int]$deploymentIntent = 1
                        Switch ($ProgramType) {
                            "Available" { 
                               $deploymentIntent = 2
                            }
                            "Required" { 
                               $deploymentIntent = 1
                            }
                        }

                        if($ProgramType -ne "LanguagePack") {
                            $comment = $ProgramType.ToString() + "-" + $ChannelName + "-" + $Bitness.ToString() + "-" + $Collection.ToString()
                            if ($CustomName) {
                               $comment += "-$CustomName" 
                            }
                        } else {
                            if($programName -like "*Multi*"){
                                $programComment = $ProgramName.Split("-")[3] + $ProgramName.Split("-")[4]
                            } else {
                                $programComment = $ProgramName.Split("-")[3]
                            }
                            $comment = $ProgramType.ToString() + "-" + $ChannelName + "-" + $Bitness.ToString() + "-$programComment-" + $Collection.ToString()
                        }

                        $packageDeploy = Get-WmiObject -Namespace "root\sms\site_$SiteCode" -Class SMS_Advertisement  | where {$_.PackageId -eq $package.PackageId -and $_.Comment -eq $comment }
                        if ($packageDeploy.Count -eq 0) {
                            try {
                                $packageId = $package.PackageId

                                if ($Program) {
                                    $ProgramName = $Program.ProgramName

     	                            Start-CMPackageDeployment -CollectionName "$Collection" -PackageId $packageId -ProgramName "$ProgramName" `
                                                                -StandardProgram  -DeployPurpose $DeploymentPurpose.ToString() -RerunBehavior AlwaysRerunProgram `
                                                                -ScheduleEvent AsSoonAsPossible -FastNetworkOption RunProgramFromDistributionPoint `
                                                                -SlowNetworkOption RunProgramFromDistributionPoint `
                                                                -AllowSharedContent $false -Comment $comment

                                    Write-Host "`tDeployment created for: $packageName ($ProgramName)"
                                    WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Deployment created for: $packageName ($ProgramName)" -LogFilePath $LogFilePath
                                } else {
                                    Write-Host "Could Not find Program in Package for Type: $ProgramType - Channel: $ChannelName" -ForegroundColor White -BackgroundColor Red
                                    WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Could Not find Program in Package for Type: $ProgramType - Channel: $ChannelName" -LogFilePath $LogFilePath
                                }
                            } catch {
                                [string]$ErrorMessage = $_.ErrorDetails 
                                if ($ErrorMessage.ToLower().Contains("Could not find property PackageID".ToLower())) {
                                    Write-Host 
                                    Write-Host "Package: $packageName"
                                    Write-Host "The package has not finished deploying to the distribution points." -BackgroundColor Red
                                    Write-Host "Please try this command against once the distribution points have been updated" -BackgroundColor Red

                                    WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Package: $packageName" -LogFilePath $LogFilePath
                                    WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "The package has not finished deploying to the distribution points." -LogFilePath $LogFilePath
                                    WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Please try this command against once the distribution points have been updated" -LogFilePath $LogFilePath
                                } else {
                                    throw
                                }
                            }  
                        } else {
                            Write-Host "`tDeployment already exists for: $packageName ($ProgramName)"
                            WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Deployment already exists for: $packageName ($ProgramName)" -LogFilePath $LogFilePath
                        }
                   } else {
                        Write-Host "Could Not find Program in Package for Type: $ProgramType - Channel: $ChannelName - Bitness: $Bitness" -ForegroundColor White -BackgroundColor Red
                        WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Could Not find Program in Package for Type: $ProgramType - Channel: $ChannelName - Bitness: $Bitness" -LogFilePath $LogFilePath
                   }
                } else {
                    WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Package does not exist: $packageName" -LogFilePath $LogFilePath
                    throw "Package does not exist: $packageName"
                }
            }
        }
       } catch {
         throw;
       }
    }
    End {
        Set-ExecutionPolicy $currentExecutionPolicy -Scope Process -Force
        Set-Location $startLocation 
    }
}

function Get-CMOfficeDistributionStatus{
Param(
    [Parameter()]
    [bool]$WaitForDistributionToFinish = $false,

    [Parameter()]
    [string]$LogFilePath
)

Begin{
    $currentExecutionPolicy = Get-ExecutionPolicy
	Set-ExecutionPolicy Unrestricted -Scope Process -Force  
    $startLocation = Get-Location
}

Process{
    $currentFileName = Get-CurrentFileName
    Set-Alias -name LINENUM -value Get-CurrentLineNumber

    $SiteCode = GetLocalSiteCode -SiteCode $SiteCode
    $Package = CheckIfPackageExists
    $PkgID = $Package.PackageID

    $Status = GetQueryStatus -SiteCode $SiteCode -PkgID $PkgID -LogFilePath $LogFilePath

    if($WaitForDistributionToFinish){
        [string[]]$trackProgress = @()
        $currentTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"            
        Do{
            $Status = GetQueryStatus -SiteCode $SiteCode -PkgID $PkgID -LogFilePath $LogFilePath

            if(!$trackProgress){
                if($Status.DateTime -ge $currentTime){
                    $trackProgress += $Status.Status
                    $updateRunning = $true
                    WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError $Status -LogFilePath $LogFilePath
                    $Status
                }
            }else{
                if($trackProgress -notcontains $Status.Status){
                    $trackProgress += $Status.Status
                    $updateRunning = $true
                    WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError $Status -LogFilePath $LogFilePath
                    $Status
                }
            }            
            
        }until($Status.Operation -eq 'In Progress')
               
        Do{
            $Status = GetQueryStatus -SiteCode $SiteCode -PkgID $PkgID -LogFilePath $LogFilePath

            if(!$trackProgress){
                $trackProgress += $Status.Status
                $Status
            }else{
                if($trackProgress -notcontains $Status.Status){
                    $trackProgress += $Status.Status
                    $updateRunning = $true
                    WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError $Status -LogFilePath $LogFilePath
                    $Status
                }
            }

            if(($Status.Operation -eq 'Failed') -or`
               ($Status.Operation -eq 'Error')){
                $trackProgress += $Status.Status
                WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError $Status -LogFilePath $LogFilePath
                $Status

                break;
            }

        }until($Status.Operation -eq 'Success')  
              
        Do{
            $Status = GetQueryStatus -SiteCode $SiteCode -PkgID $PkgID -LogFilePath $LogFilePath
        
            if($Status.Operation -eq "Success"){
                if($Status.Status -eq 'Content was distributed to distribution point'){
                    $trackProgress += $Status.Status
                    $allComplete = $true
                    $updateRunning = $false
                    WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError $Status -LogFilePath $LogFilePath
                    $Status
                }
                else{        
                    if($trackProgress -notcontains $Status.Status){
                        $trackProgress += $Status.Status
                        $updateRunning = $true
                        WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError $Status -LogFilePath $LogFilePath
                        $Status
                    }
                }
            }

            if($Status.Operation -eq "Failed"){
                $updateRunning = $false
                WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError $Status -LogFilePath $LogFilePath
                $Status
            }  
            
        }while($updateRunning -eq $true)
    }else{
        WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError $Status -LogFilePath $LogFilePath
        $Status
    }
}

End{
    Set-ExecutionPolicy $currentExecutionPolicy -Scope Process -Force
    Set-Location $startLocation
}

}

function UpdateConfigurationXml() {
    [CmdletBinding()]	
    Param
	(
		[Parameter(Mandatory=$true)]
		[String]$Path = "",

		[Parameter(Mandatory=$true)]
		[String]$Channel = "",

		[Parameter(Mandatory=$true)]
		[String]$Bitness,

		[Parameter()]
		[String]$SourcePath = $NULL,
		
        [Parameter()]
		[String]$Language,

        [Parameter()]
        [string]$LogFilePath
        
	) 
    Process {
      $currentFileName = Get-CurrentFileName
      Set-Alias -name LINENUM -value Get-CurrentLineNumber

      switch($channel){
        "MonthlyTargeted"{
            $Channel = "Insiders"
        }
        "Monthly"{
            $Channel = "Monthly"
        }
        "SemiAnnualTargeted"{
            $Channel = "Targeted"
        }
        "SemiAnnual"{
            $Channel = "Broad"
        }
      }

	  $doc = [Xml] (Get-Content $Path)

      $addNode = $doc.Configuration.Add
      $languageNode = $addNode.Product.Language

      if ($addNode.OfficeClientEdition) {
          $addNode.OfficeClientEdition = $Bitness
      } else {
          $addNode.SetAttribute("OfficeClientEdition", $Bitness)
      }

      if ($addNode.Channel) {
          $addNode.Channel = $Channel
      } else {
          $addNode.SetAttribute("Channel", $Channel)
      }

      if ($addNode.SourcePath) {
          $addNode.SourcePath = $SourcePath
      } else {
          $addNode.SetAttribute("SourcePath", $SourcePath)
      }

      if($Language){
          if ($languageNode.ID){
              if($languageNode.ID -contains $Language) {
                  Write-Host "$Language already exists in the xml"
                  WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "$Language already exists in the xml" -LogFilePath $LogFilePath
              } else {
                  $newLanguageElement = $doc.CreateElement("Language")
                  $newLanguage = $doc.Configuration.Add.Product.AppendChild($newLanguageElement)
                  $newLanguage.SetAttribute("ID", $Language)
              }
          } else {
              $languageNode.SetAttribute("ID", $language)
          }
     }

      $doc.Save($Path)
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

function CheckIfPackageExists() {
    [CmdletBinding()]	
    Param
	(

    )
    Begin
    {
        $startLocation = Get-Location
    }
    Process {
       LoadCMPrereqs

       $packageName = "Office 365 ProPlus"

       $existingPackage = Get-CMPackage | Where { $_.Name -eq $packageName }
       if ($existingPackage) {
         return $existingPackage
       }

       return $null
    }
}

function CheckIfVersionExists() {
    [CmdletBinding()]	
    Param
	(
	   [Parameter(Mandatory=$True)]
	   [String]$Version,

		[Parameter()]
		[String]$Channel
    )
    Begin
    {
        $startLocation = Get-Location
    }
    Process {
       LoadCMPrereqs

       $VersionName = "$Channel - $Version"

       $packageName = "Office 365 ProPlus"

       $existingPackage = Get-CMPackage | Where { $_.Name -eq $packageName -and $_.Version -eq $Version }
       if ($existingPackage) {
         return $true
       }

       return $false
    }
}

function LoadCMPrereqs() {
    [CmdletBinding()]	
    Param
	(
	    [Parameter()]
	    [String]$SiteCode = $null,

	    [Parameter()]
	    [String]$CMPSModulePath = $NULL
    )
    Begin
    {
        $currentExecutionPolicy = Get-ExecutionPolicy
	    Set-ExecutionPolicy Unrestricted -Scope Process -Force  
        $startLocation = Get-Location
    }
    Process {

        $CMModulePath = GetCMPSModulePath -CMPSModulePath $CMPSModulePath -LogFilePath $LogFilePath
    
        if ($CMModulePath) {
            Import-Module $CMModulePath

            if (!$SiteCode) {
               $SiteCode = (Get-ItemProperty -Path "hklm:\SOFTWARE\Microsoft\SMS\Identification" -Name "Site Code").'Site Code'
            }

            Set-Location "$SiteCode`:"	
        }
    }
}

function GetLocalSiteCode() {
    [CmdletBinding()]	
    Param
	(
	    [Parameter()]
	    [String]$SiteCode = $null
    )
    Begin
    {

    }
    Process {
        if (!$SiteCode) {
            $SiteCode = (Get-ItemProperty -Path "hklm:\SOFTWARE\Microsoft\SMS\Identification" -Name "Site Code").'Site Code'
        }
        return $SiteCode
    }
}

function CreateCMPackage() {
    [CmdletBinding()]	
    Param
	(
		[Parameter()]
		[String]$Name = "Office ProPlus Deployment",
		
		[Parameter(Mandatory=$True)]
		[String]$Path,

		[Parameter()]
		[String]$Version,

		[Parameter()]
		[String]$Channel,

		[Parameter()]
		[String]$CustomPackageShareName = $null,

		[Parameter()]	
		[Bool]$UpdateOnlyChangedBits = $true,

        [Parameter()]
        [string]$LogFilePath
	) 
    $currentFileName = Get-CurrentFileName
    Set-Alias -name LINENUM -value Get-CurrentLineNumber

    $package = Get-CMPackage | Where { $_.Name -eq $Name }
    if($package -eq $null -or !$package)
    {
        Write-Host "`tCreating Package: $Name"
        WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Creating Package: $Name" -LogFilePath $LogFilePath
        $package = New-CMPackage -Name $Name -Path $path -Version $Version
    } else {
        Write-Host "`t`tPackage Already Exists: $Name"  
        WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Package Already Exists: $Name" -LogFilePath $LogFilePath      
    }
		
    Write-Host "`t`tSetting Package Properties" 
    WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Setting Package Properties" -LogFilePath $LogFilePath     

    $VersionName = "$Channel - $Version"

    if ($CustomPackageShareName) {
	    Set-CMPackage -Id $package.PackageId -Priority Normal -EnableBinaryDeltaReplication $UpdateOnlyChangedBits `
                      -CopyToPackageShareOnDistributionPoint $True -Version $Version -CustomPackageShareName $CustomPackageShareName
    } else {
	    Set-CMPackage -Id $package.PackageId -Priority Normal -EnableBinaryDeltaReplication $UpdateOnlyChangedBits `
                      -CopyToPackageShareOnDistributionPoint $True -Version $Version
    }

    $package = Get-CMPackage | Where { $_.Name -eq $Name -and $_.Version -eq $Version }
    return $package
}

function RemovePreviousCMPackages() {
    [CmdletBinding()]	
    Param
	(
		[Parameter()]
		[String]$Name = "Office ProPlus Deployment",
		
		[Parameter()]
		[String]$Version,

        [Parameter()]
        [string]$LogFilePath      
	) 
    $currentFileName = Get-CurrentFileName
    Set-Alias -name LINENUM -value Get-CurrentLineNumber

    if ($Version) {
        $packages = Get-CMPackage | Where { $_.Name -eq $Name -and $_.Version -ne $Version }
        foreach ($package in $packages) {
           $packageName = $package.Name
           $pkversion = $package.Version

           Write-Host "Removing previous version: $packageName - $pkversion"
           WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Removing previous version: $packageName - $pkversion" -LogFilePath $LogFilePath       
           Remove-CMPackage -Id $package.PackageId -Force | Out-Null
        }
    }
}

function CreateCMProgram() {
    [CmdletBinding()]	
    Param
	(
		[Parameter(Mandatory=$True)]
		[String]$PackageID,
		
		[Parameter(Mandatory=$True)]
		[String]$CommandLine, 

		[Parameter(Mandatory=$True)]
		[String]$Name,
		
		[Parameter(Mandatory=$True)]
		[String]$Comment = $null,

		[Parameter()]
		[String[]] $RequiredPlatformNames = @(),

        [Parameter()]
        [string]$LogFilePath 

	) 
    $currentFileName = Get-CurrentFileName
    Set-Alias -name LINENUM -value Get-CurrentLineNumber

    $program = Get-CMProgram | Where { $_.PackageID -eq $PackageID -and $_.Comment -eq $Comment -and $_.ProgramName -eq $Name }

    if($program -eq $null -or !$program) {
        Write-Host "`t`tCreating Program: $Name ..."	   
        WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Creating Program: $Name ..." -LogFilePath $LogFilePath     
	    $program = New-CMProgram -PackageId $PackageID -StandardProgramName $Name -DriveMode RenameWithUnc `
                                 -CommandLine $CommandLine -ProgramRunType OnlyWhenUserIsLoggedOn `
                                 -RunMode RunWithAdministrativeRights -UserInteraction $true -RunType Normal
    } else {
        Write-Host "`t`tProgram Already Exists: $Name"
        WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Program Already Exists: $Name" -LogFilePath $LogFilePath     
    }

    if ($program) {
        Set-CMProgram -InputObject $program -Comment $Comment -StandardProgramName $Name -StandardProgram
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

function GetSupportedPlatforms([String[]] $requiredPlatformNames){
    $computerName = $env:COMPUTERNAME
    #$assignedSite = $([WmiClass]"\\$computerName\ROOT\ccm:SMS_Client").getassignedsite()
    $siteCode = Get-Site -LogFilePath $LogFilePath
    $filteredPlatforms = Get-WmiObject -ComputerName $computerName -Class SMS_SupportedPlatforms -Namespace "root\sms\site_$siteCode" | Where-Object {$_.IsSupported -eq $true -and  $_.OSName -like 'Win NT' -and ($_.OSMinVersion -match "6\.[0-9]{1,2}\.[0-9]{1,4}\.[0-9]{1,4}" -or $_.OSMinVersion -match "10\.[0-9]{1,2}\.[0-9]{1,4}\.[0-9]{1,4}") -and ($_.OSPlatform -like 'I386' -or $_.OSPlatform -like 'x64')}

    $requiredPlatforms = $filteredPlatforms| Where-Object {$requiredPlatformNames.Contains($_.DisplayText) } #| Select DisplayText, OSMaxVersion, OSMinVersion, OSName, OSPlatform | Out-GridView

    $supportedPlatforms = @()

    foreach($p in $requiredPlatforms)
    {
        $osDetail = ([WmiClass]("\\$computerName\root\sms\site_$siteCode`:SMS_OS_Details")).CreateInstance()    
        $osDetail.MaxVersion = $p.OSMaxVersion
        $osDetail.MinVersion = $p.OSMinVersion
        $osDetail.Name = $p.OSName
        $osDetail.Platform = $p.OSPlatform

        $supportedPlatforms += $osDetail
    }

    $supportedPlatforms
}

function CreateDownloadXmlFile([string]$Path, [string]$ConfigFileName){
	#1 - Set the correct version number to update Source location
	$sourceFilePath = "$path\$configFileName"
    $localSourceFilePath = ".\$configFileName"

    Set-Location $PSScriptRoot

    if (Test-Path -Path $localSourceFilePath) {   
	  $doc = [Xml] (Get-Content $localSourceFilePath)

      $addNode = $doc.Configuration.Add
	  $addNode.OfficeClientEdition = $bitness

      $doc.Save($sourceFilePath)
    } else {
      $doc = New-Object System.XML.XMLDocument

      $configuration = $doc.CreateElement("Configuration");
      $a = $doc.AppendChild($configuration);

      $addNode = $doc.CreateElement("Add");
      $addNode.SetAttribute("OfficeClientEdition", $bitness)
      if ($Version) {
         if ($Version.Length -gt 0) {
             $addNode.SetAttribute("Version", $Version)
         }
      }
      $a = $doc.DocumentElement.AppendChild($addNode);

      $addProduct = $doc.CreateElement("Product");
      $addProduct.SetAttribute("ID", "O365ProPlusRetail")
      $a = $addNode.AppendChild($addProduct);

      $addLanguage = $doc.CreateElement("Language");
      $addLanguage.SetAttribute("ID", "en-us")
      $a = $addProduct.AppendChild($addLanguage);

	  $doc.Save($sourceFilePath)
    }
}

function CreateUpdateXmlFile([string]$Path, [string]$ConfigFileName, [string]$Bitness, [string]$Version){
    $newConfigFileName = $ConfigFileName -replace '\.xml'
    $newConfigFileName = $newConfigFileName + "$Bitness" + ".xml"

    Copy-Item -Path ".\$ConfigFileName" -Destination ".\$newConfigFileName"
    $ConfigFileName = $newConfigFileName

    $testGroupFilePath = "$path\$ConfigFileName"
    $localtestGroupFilePath = ".\$ConfigFileName"

	$testGroupConfigContent = [Xml] (Get-Content $localtestGroupFilePath)

	$addNode = $testGroupConfigContent.Configuration.Add
	$addNode.OfficeClientEdition = $bitness
    $addNode.SourcePath = $path	

	$updatesNode = $testGroupConfigContent.Configuration.Updates
	$updatesNode.UpdatePath = $path
	$updatesNode.TargetVersion = $version

	$testGroupConfigContent.Save($testGroupFilePath)
    return $ConfigFileName
}
 
function GetCMPSModulePath() {
    [CmdletBinding()]	
    Param
	(
		[Parameter()]
		[String]$CMPSModulePath = $NULL,

        [Parameter()]
        [string]$LogFilePath
	)
    $currentFileName = Get-CurrentFileName
    Set-Alias -name LINENUM -value Get-CurrentLineNumber

    [bool]$pathExists = $false

    if ($CMPSModulePath) {
       if ($CMPSModulePath.ToLower().EndsWith(".psd1")) {
         $CMModulePath = $CMPSModulePath
         $pathExists = Test-Path -Path $CMModulePath
       }
    }

    if (!$pathExists) {
        $uiInstallDir = (Get-ItemProperty -Path "hklm:\SOFTWARE\Microsoft\SMS\Setup" -Name "UI Installation Directory").'UI Installation Directory'
        $CMModulePath = Join-Path $uiInstallDir "bin\ConfigurationManager.psd1"

        $pathExists = Test-Path -Path $CMModulePath
        if (!$pathExists) {
            $CMModulePath = "$env:ProgramFiles\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1"
            $pathExists = Test-Path -Path $CMModulePath
        }
    }

    if (!$pathExists) {
       $uiAdminPath = ${env:SMS_ADMIN_UI_PATH}
       if ($uiAdminPath.ToLower().EndsWith("\bin")) {
           $dirInfo = $uiAdminPath
       } else {
           $dirInfo = ([System.IO.DirectoryInfo]$uiAdminPath).Parent.FullName
       }
      
       $CMModulePath = $dirInfo + "\ConfigurationManager.psd1"
       $pathExists = Test-Path -Path $CMModulePath
    }

    if (!$pathExists) {
       $CMModulePath = "${env:ProgramFiles(x86)}\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1"
       $pathExists = Test-Path -Path $CMModulePath
    }

    if (!$pathExists) {
       $CMModulePath = "${env:ProgramFiles(x86)}\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1"
       $pathExists = Test-Path -Path $CMModulePath
    }

    if (!$pathExists) {
        WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Cannot find the ConfigurationManager.psd1 file. Please use the -CMPSModulePath parameter to specify the location of the PowerShell Module" -LogFilePath $LogFilePath
        throw "Cannot find the ConfigurationManager.psd1 file. Please use the -CMPSModulePath parameter to specify the location of the PowerShell Module"
    }

    return $CMModulePath
}

function Get-Site {
Param(
    [Parameter()]
    [string[]]$computerName = $env:COMPUTERNAME,

    [Parameter()]
    [string]$LogFilePath
)
    $currentFileName = Get-CurrentFileName
    Set-Alias -name LINENUM -value Get-CurrentLineNumber

    Get-WmiObject -ComputerName $ComputerName -Namespace "root\SMS" -Class "SMS_ProviderLocation" | foreach-object{ 
        if ($_.ProviderForLocalSite -eq $true){$SiteCode=$_.sitecode} 
    } 
    if ($SiteCode -eq "") { 
        WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Sitecode of ConfigMgr Site at " + $ComputerName + " could not be determined." -LogFilePath $LogFilePath
        throw ("Sitecode of ConfigMgr Site at " + $ComputerName + " could not be determined.") 
    } else { 
        Return $SiteCode 
    } 
}

function DownloadBits() {
    [CmdletBinding()]	
    Param
	(
	    [Parameter()]
	    [OfficeBranch]$Branch = $null
	)

    $DownloadScript = "$PSScriptRoot\Download-OfficeProPlusBranch.ps1"
    if (Test-Path -Path $DownloadScript) {
      
    }
}

Function GetScriptRoot() {
 process {
     [string]$scriptPath = "."

     if ($PSScriptRoot) {
       $scriptPath = $PSScriptRoot
     } else {
       $scriptPath = Split-Path -parent $MyInvocation.MyCommand.Definition
     }

     return $scriptPath
 }
}

function GetQueryStatus(){
Param(
    [Parameter()]
    [string]$SiteCode,

    [Parameter()]
    [string]$PkgID,

    [Parameter()]
    [string]$LogFilePath
)
    $currentFileName = Get-CurrentFileName
    Set-Alias -name LINENUM -value Get-CurrentLineNumber

    $query = Get-WmiObject -NameSpace Root\SMS\Site_$SiteCode -Class SMS_DistributionDPStatus -Filter "PackageID='$PkgID'" | Select Name, MessageID, MessageState, LastUpdateDate

    if ($query -eq $null)
    {  
        WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "PackageID not found" -LogFilePath $LogFilePath
        throw "PackageID not found"
    }

    foreach ($objItem in $query){

        $DPName = $objItem.Name
        $UpdDate = [System.Management.ManagementDateTimeconverter]::ToDateTime($objItem.LastUpdateDate)

        switch ($objItem.MessageState)
        {
            1         {$Status = "Success"}
            2         {$Status = "In Progress"}
            3         {$Status = "Failed"}
            4         {$Status = "Error"}
        }

        switch ($objItem.MessageID)
        {
            2300      {$Message = "Content is beginning to process"}
            2301      {$Message = "Content has been processed successfully"}
            2303      {$Message = "Failed to process package"}
            2311      {$Message = "Distribution Manager has successfully created or updated the package"}
            2303      {$Message = "Content was successfully refreshed"}
            2323      {$Message = "Failed to initialize NAL"}
            2324      {$Message = "Failed to access or create the content share"}
            2330      {$Message = "Content was distributed to distribution point"}
            2342      {$Message = "Content is beginning to distribute"}
            2354      {$Message = "Failed to validate content status file"}
            2357      {$Message = "Content transfer manager was instructed to send content to Distribution Point"}
            2360      {$Message = "Status message 2360 unknown"}
            2370      {$Message = "Failed to install distribution point"}
            2371      {$Message = "Waiting for prestaged content"}
            2372      {$Message = "Waiting for content"}
            2376      {$Message = "Distribution Manager created a snapshot for content"}
            2380      {$Message = "Content evaluation has started"}
            2381      {$Message = "An evaluation task is running. Content was added to Queue"}
            2382      {$Message = "Content hash is invalid"}
            2383      {$Message = "Failed to validate content hash"}
            2384      {$Message = "Content hash has been successfully verified"}
            2391      {$Message = "Failed to connect to remote distribution point"}
            2397      {$Message = "Detail will be available after the server finishes processing the messages"}
            2398      {$Message = "Content Status not found"}
            8203      {$Message = "Failed to update package"}
            8204      {$Message = "Content is being distributed to the distribution Point"}
            8211      {$Message = "Failed to update package"}
        }

        $Displayvalue = showTaskStatus -Operation $Status -Status $Message -DateTime $UpdDate

    }

    return $Displayvalue
}

function showTaskStatus() {
    [CmdletBinding()]
    Param(
        [Parameter()]
        [string] $Operation = "",

        [Parameter()]
        [string] $Status = "",

        [Parameter()]
        [string] $DateTime = ""
    )

    $Result = New-Object -TypeName PSObject 
    Add-Member -InputObject $Result -MemberType NoteProperty -Name "Operation" -Value $Operation
    Add-Member -InputObject $Result -MemberType NoteProperty -Name "Status" -Value $Status
    Add-Member -InputObject $Result -MemberType NoteProperty -Name "DateTime" -Value $DateTime
    return $Result
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

$currentFileName = Get-CurrentFileName
Set-Alias -name LINENUM -value Get-CurrentLineNumber

$scriptPath = GetScriptRoot

$shareFunctionsPath = "$scriptPath\SharedFunctions.ps1"
if ($scriptPath.StartsWith("\\")) {
} else {
    if (!(Test-Path -Path $shareFunctionsPath)) {
        WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Missing Dependency File SharedFunctions.ps1" -LogFilePath $LogFilePath    
        throw "Missing Dependency File SharedFunctions.ps1"    
    }
}
. $shareFunctionsPath

