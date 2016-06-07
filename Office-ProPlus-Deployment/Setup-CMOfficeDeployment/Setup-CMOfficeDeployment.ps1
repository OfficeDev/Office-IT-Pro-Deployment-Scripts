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
          Deferred = 3
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
        UpdateWithTask = 5
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
        [OfficeChannel[]] $Channels = @(1,2,3),

        [Parameter(Mandatory=$true)]
	    [String]$OfficeFilesPath = $NULL,

        [Parameter()]
        [ValidateSet("en-us","ar-sa","bg-bg","zh-cn","zh-tw","hr-hr","cs-cz","da-dk","nl-nl","et-ee","fi-fi","fr-fr","de-de","el-gr","he-il","hi-in","hu-hu","id-id","it-it",
                    "ja-jp","kk-kh","ko-kr","lv-lv","lt-lt","ms-my","nb-no","pl-pl","pt-br","pt-pt","ro-ro","ru-ru","sr-latn-rs","sk-sk","sl-si","es-es","sv-se","th-th",
                    "tr-tr","uk-ua")]
        [string[]] $Languages = ("en-us"),

        [Parameter()]
        [Bitness] $Bitness = 0,

        [Parameter()]
        [string] $Version = $NULL
        
    )

    Process {
       if (Test-Path "$PSScriptRoot\Download-OfficeProPlusChannels.ps1") {
         . "$PSScriptRoot\Download-OfficeProPlusChannels.ps1"
       } else {
         throw "Dependency file missing: $PSScriptRoot\Download-OfficeProPlusChannels.ps1"
       }

       $ChannelList = @("FirstReleaseCurrent", "Current", "FirstReleaseDeferred", "Deferred")
       $ChannelXml = Get-ChannelXml -FolderPath $OfficeFilesPath -OverWrite $true

       foreach ($Channel in $ChannelList) {
         if ($Channels -contains $Channel) {

            $selectChannel = $ChannelXml.UpdateFiles.baseURL | Where {$_.branch -eq $Channel.ToString() }
            $latestVersion = Get-ChannelLatestVersion -ChannelUrl $selectChannel.URL -Channel $Channel
            $ChannelShortName = ConvertChannelNameToShortName -ChannelName $Channel

            if ($Version) {
               $latestVersion = $Version
            }

            Download-OfficeProPlusChannels -TargetDirectory $OfficeFilesPath  -Channels $Channel -Version $latestVersion -UseChannelFolderShortName $true -Languages $Languages -Bitness $Bitness

            $cabFilePath = "$env:TEMP/ofl.cab"
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
        [OfficeChannel[]] $Channels = @(1,2,3),

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
	    [String]$CMPSModulePath = $NULL
    )
    Begin
    {
        $currentExecutionPolicy = Get-ExecutionPolicy
	    Set-ExecutionPolicy Unrestricted -Scope Process -Force  
        $startLocation = Get-Location
    }
    Process {
       try {

       Check-AdminAccess

       $cabFilePath = "$OfficeSourceFilesPath\ofl.cab"
       if (Test-Path $cabFilePath) {
            Copy-Item -Path $cabFilePath -Destination "$PSScriptRoot\ofl.cab" -Force
       }

       $ChannelList = @("FirstReleaseCurrent", "Current", "FirstReleaseDeferred", "Deferred")
       $ChannelXml = Get-ChannelXml -FolderPath $OfficeSourceFilesPath -OverWrite $false

       foreach ($Channel in $ChannelList) {
         if ($Channels -contains $Channel) {
           $selectChannel = $ChannelXml.UpdateFiles.baseURL | Where {$_.branch -eq $Channel.ToString() }
           $latestVersion = Get-ChannelLatestVersion -ChannelUrl $selectChannel.URL -Channel $Channel -FolderPath $OfficeFilesPath -OverWrite $false

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
                    throw "Channel Folder Missing: $officeFileChannelPath - Ensure that you have downloaded the Channel you are trying to deploy"
                }

                [System.IO.Directory]::CreateDirectory($officeFileTargetPath) | Out-Null

                if ($MoveSourceFiles) {
                    Move-Item -Path $officeFileChannelPath -Destination $officeFileTargetPath -Force
                } else {
                    Copy-Item -Path $officeFileChannelPath -Destination $officeFileTargetPath -Recurse -Force
                }

                $cabFilePath = "$OfficeSourceFilesPath\ofl.cab"
                if (Test-Path $cabFilePath) {
                    Copy-Item -Path $cabFilePath -Destination "$LocalPath\ofl.cab" -Force
                }
           } else {
              if (Test-Path -Path "$LocalChannelPath\Office") {
                 Remove-Item -Path "$LocalChannelPath\Office" -Force -Recurse
              }
           }

           $cabFilePath = "$env:TEMP/ofl.cab"
           if (!(Test-Path $cabFilePath)) {
                Copy-Item -Path "$LocalPath\ofl.cab" -Destination $cabFilePath -Force
           }

           CreateMainCabFiles -LocalPath $LocalPath -ChannelShortName $ChannelShortName -LatestVersion $latestVersion

           $DeploymentFilePath = "$PSSCriptRoot\DeploymentFiles\*.*"
           if (Test-Path -Path $DeploymentFilePath) {
             Copy-Item -Path $DeploymentFilePath -Destination "$LocalPath" -Force -Recurse
           } else {
             throw "Deployment folder missing: $DeploymentFilePath"
           }

           LoadCMPrereqs -SiteCode $SiteCode -CMPSModulePath $CMPSModulePath

           if (!($existingPackage)) {
              $package = CreateCMPackage -Name $packageName -Path $Path -Channel $Channel -UpdateOnlyChangedBits $UpdateOnlyChangedBits -CustomPackageShareName $CustomPackageShareName
           } else {
              Write-Host "`tPackage Already Exists: $packageName"
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
        [OfficeChannel[]] $Channels = @(1,2,3),

        [Parameter()]
	    [String]$OfficeSourceFilesPath = $NULL,

        [Parameter()]
	    [bool]$MoveSourceFiles = $false,

	    [Parameter()]
	    [String]$SiteCode = $null,

	    [Parameter()]
	    [String]$CMPSModulePath = $NULL,

        [Parameter()]
	    [bool]$UpdateDistributionPoints = $true
    )
    Begin
    {
        $currentExecutionPolicy = Get-ExecutionPolicy
	    Set-ExecutionPolicy Unrestricted -Scope Process -Force  
        $startLocation = Get-Location
    }
    Process {
       try {

       Check-AdminAccess

       $cabFilePath = "$OfficeSourceFilesPath\ofl.cab"
       if (Test-Path $cabFilePath) {
            Copy-Item -Path $cabFilePath -Destination "$PSScriptRoot\ofl.cab" -Force
       }

       $ChannelList = @("FirstReleaseCurrent", "Current", "FirstReleaseDeferred", "Deferred")
       $ChannelXml = Get-ChannelXml -FolderPath $OfficeSourceFilesPath -OverWrite $false

       foreach ($Channel in $ChannelList) {
         if ($Channels -contains $Channel) {
           $selectChannel = $ChannelXml.UpdateFiles.baseURL | Where {$_.branch -eq $Channel.ToString() }
           $latestVersion = Get-ChannelLatestVersion -ChannelUrl $selectChannel.URL -Channel $Channel -FolderPath $OfficeFilesPath -OverWrite $false

           $ChannelShortName = ConvertChannelNameToShortName -ChannelName $Channel
           $existingPackage = CheckIfPackageExists
           if (!($existingPackage)) {
              throw "No Package Exists to Update. Please run the Create-CMOfficePackage function first to create the package."
           }

           $packagePath = $existingPackage.PkgSourcePath
           if ($packagePath.StartsWith("\\")) {
               $shareName = $packagePath.Split("\")[3]
           }

           $existingShare = Get-Fileshare -Name $shareName
           if (!($existingShare)) {
              throw "No Package Exists to Update. Please run the Create-CMOfficePackage function first to create the package."
           }

           $packageName = $existingPackage.Name
           $packageId = $existingPackage.PackageID

           Write-Host "`tUpdating Package: $packageName"

           $Path = $existingPackage.PkgSourcePath

           $packageName = "Office 365 ProPlus"
           $ChannelPath = "$Path\$Channel"
           $LocalPath = $existingShare.Path
           $LocalChannelPath = $existingShare.Path + "\SourceFiles"

           [System.IO.Directory]::CreateDirectory($LocalChannelPath) | Out-Null
                          
           if ($OfficeSourceFilesPath) {
                Write-Host "`t`tUpdating Source Files..."

                $officeFileChannelPath = "$OfficeSourceFilesPath\$ChannelShortName"
                $officeFileTargetPath = "$LocalChannelPath\$Channel"

                if (!(Test-Path -Path $officeFileChannelPath)) {
                    throw "Channel Folder Missing: $officeFileChannelPath - Ensure that you have downloaded the Channel you are trying to deploy"
                }

                [System.IO.Directory]::CreateDirectory($officeFileTargetPath) | Out-Null

                if ($MoveSourceFiles) {
                    Move-Item -Path $officeFileChannelPath -Destination $LocalChannelPath -Force
                } else {
                    Copy-Item -Path $officeFileChannelPath -Destination $LocalChannelPath -Recurse -Force
                }

                $cabFilePath = "$OfficeSourceFilesPath\ofl.cab"
                if (Test-Path $cabFilePath) {
                    Copy-Item -Path $cabFilePath -Destination "$LocalPath\ofl.cab" -Force
                }
           }

           $cabFilePath = "$env:TEMP/ofl.cab"
           if (!(Test-Path $cabFilePath)) {
                Copy-Item -Path "$LocalPath\ofl.cab" -Destination $cabFilePath -Force
           }

           CreateMainCabFiles -LocalPath $LocalPath -ChannelShortName $ChannelShortName -LatestVersion $latestVersion

           $DeploymentFilePath = "$PSSCriptRoot\DeploymentFiles\*.*"
           if (Test-Path -Path $DeploymentFilePath) {
             Write-Host "`t`tUpdating Deployment Files..."
             Copy-Item -Path $DeploymentFilePath -Destination "$LocalPath" -Force -Recurse
           } else {
             throw "Deployment folder missing: $DeploymentFilePath"
           }

           LoadCMPrereqs -SiteCode $SiteCode -CMPSModulePath $CMPSModulePath

           if ($UpdateDistributionPoints) {
              Write-Host "`t`tUpdating Distribution Points..."
              Update-CMDistributionPoint -PackageId $packageId
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
        [OfficeChannel[]] $Channels = @(1,2,3),

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
	    [String]$CustomName = $NULL
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
             throw "CustomName is too long.  Must be less then 50 Characters"
         }

         foreach ($channel in $Channels) {
             LoadCMPrereqs -SiteCode $SiteCode -CMPSModulePath $CMPSModulePath

             $LargeDrv = Get-LargestDrive
             $LocalPath = "$LargeDrv\OfficeDeployment"

             $channelShortName = ConvertChannelNameToShortName -ChannelName $channel

             $existingPackage = CheckIfPackageExists
             if (!($existingPackage)) {
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
                        throw "Configuration file does not exist: $ConfigurationXml"
                     }

                     #[guid]::NewGuid().Guid.ToString()

                     $configId = "Config-$channel-$platform-Bit"
                     $configFileName = $configId + ".xml"

                     if ($CustomName) {
                         $configFileName = $configId + "-" + $CustomName + ".xml"
                     }

                     $configFilePath = "$LocalPath\$configFileName"

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

                     UpdateConfigurationXml -Path $configFilePath -Channel $channel -Bitness $platform -SourcePath $sourcePath

                     $ProgramName = "Deploy $channelShortNameLabel Channel with Config File - $platform-Bit"

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

                    CreateCMProgram -Name $ProgramName -PackageID $packageId -RequiredPlatformNames $requiredPlatformNames -CommandLine $CommandLine -Comment $comment
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
        [OfficeChannel[]] $Channels = @(1,2,3),

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
    Process 
    {
       try {

         Check-AdminAccess

         LoadCMPrereqs -SiteCode $SiteCode -CMPSModulePath $CMPSModulePath

         $existingPackage = CheckIfPackageExists
         if (!($existingPackage)) {
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
                throw "Required file missing: $OSSourcePath"
             } else {
                 if (!(Test-ItemPathUNC -Path $SharePath -FileName "Change-OfficeChannel.ps1")) {
                    Copy-ItemUNC -SourcePath $OSSourcePath -TargetPath $SharePath -FileName "Change-OfficeChannel.ps1"
                 }

                 [string]$packageId = $existingPackage.PackageId
                 if ($packageId) {
                    $comment = "ChangeChannel-$channel"

                    CreateCMProgram -Name $ProgramName -PackageID $packageId -CommandLine $CommandLine -RequiredPlatformNames $requiredPlatformNames -Comment $comment
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
	    [String]$CMPSModulePath = $NULL
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

         Check-AdminAccess

         LoadCMPrereqs -SiteCode $SiteCode -CMPSModulePath $CMPSModulePath

         $existingPackage = CheckIfPackageExists
         if (!($existingPackage)) {
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
            throw "Required file missing: $OSSourcePath"
         } else {
             if (!(Test-ItemPathUNC -Path $SharePath -FileName "Change-OfficeChannel.ps1")) {
                Copy-ItemUNC -SourcePath $OSSourcePath -TargetPath $SharePath -FileName "Change-OfficeChannel.ps1"
             }

             [string]$packageId = $existingPackage.PackageId
             if ($packageId) {
                $comment = "RollBack"

                CreateCMProgram -Name $ProgramName -PackageID $packageId -CommandLine $CommandLine -RequiredPlatformNames $requiredPlatformNames -Comment $comment
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
	    [bool]$UseScriptLocationAsUpdateSource = $true
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

         Check-AdminAccess

         LoadCMPrereqs -SiteCode $SiteCode -CMPSModulePath $CMPSModulePath

         $existingPackage = CheckIfPackageExists
         if (!($existingPackage)) {
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
            throw "Required file missing: $OSSourcePath"
         } else {
             if (!(Test-ItemPathUNC -Path $SharePath -FileName "Update-Office365Anywhere.ps1")) {
                Copy-ItemUNC -SourcePath $OSSourcePath -TargetPath $SharePath -FileName "Update-Office365Anywhere.ps1"
             }

             [string]$packageId = $existingPackage.PackageId
             if ($packageId) {
                $comment = "UpdateWithConfigMgr"

                CreateCMProgram -Name $ProgramName -PackageID $packageId -CommandLine $CommandLine -RequiredPlatformNames $requiredPlatformNames -Comment $comment
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
        [string] $StartTime = "12:00",

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
	    [bool]$UseScriptLocationAsUpdateSource = $true
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

         Check-AdminAccess

         LoadCMPrereqs -SiteCode $SiteCode -CMPSModulePath $CMPSModulePath

         $existingPackage = CheckIfPackageExists
         if (!($existingPackage)) {
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

         if ($UpdateToVersion) {
             $CommandLine += " -UpdateToVersion " + $UpdateToVersion
         }

         $SharePath = $existingPackage.PkgSourcePath

         $OSSourcePath = "$PSScriptRoot\DeploymentFiles\Update-Office365Anywhere.ps1"
         $OCScriptPath = "$SharePath\Update-Office365Anywhere.ps1"

         $OSSourcePathTask = "$PSScriptRoot\DeploymentFiles\Create-Office365AnywhereTask.ps1"
         $OCScriptPathTask = "$SharePath\Create-Office365AnywhereTask.ps1"

         if (!(Test-Path $OSSourcePath)) {
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

                CreateCMProgram -Name $ProgramName -PackageID $packageId -CommandLine $CommandLine -RequiredPlatformNames $requiredPlatformNames -Comment $comment
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
        [OfficeChannel[]] $Channels = @(1,2,3),

	    [Parameter()]
	    [string]$DistributionPoint,

	    [Parameter()]
	    [string]$DistributionPointGroupName,

	    [Parameter()]
	    [uint16]$DeploymentExpiryDurationInDays = 15,

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
    Process
    {
       try {

        Check-AdminAccess

        $package = CheckIfPackageExists

        if (!($package)) {
            throw "You must run the Create-CMOfficePackage function before running this function"
        }

        LoadCMPrereqs -SiteCode $SiteCode -CMPSModulePath $CMPSModulePath

        if ($package) {
            [string]$packageName = $package.Name

            if ($DistributionPointGroupName) {
                Write-Host "Starting Content Distribution for package: $packageName"
	            Start-CMContentDistribution -PackageName $packageName -DistributionPointGroupName $DistributionPointGroupName
            }

            if ($DistributionPoint) {
                Write-Host "Starting Content Distribution for package: $packageName"
                Start-CMContentDistribution -PackageName $packageName -DistributionPointName $DistributionPoint
            }
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
        [OfficeChannel]$Channel = "Deferred",

        [Parameter()]
	    [BitnessOptions]$Bitness = "v32",
    
	    [Parameter()]
	    [String]$SiteCode = $NULL,

	    [Parameter()]
	    [String]$CMPSModulePath = $NULL,

    	[Parameter()]
	    [DeploymentPurpose]$DeploymentPurpose = "Default",

	    [Parameter()]
	    [String]$CustomName = $NULL
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

        Check-AdminAccess

        if ($CustomName) {
           if (!($CustomName.ToLower().Contains("deploy"))) {
              $CustomName = "Deploy " + $CustomName
           }
        }

        $CustomName = $CustomName -replace ' ', ''

        if ($CustomName.Length -gt 50) {
            throw "CustomName is too long.  Must be less then 50 Characters"
        }

        $strBitness = $Bitness.ToString() -Replace "v", ""

        $ChannelList = @("FirstReleaseCurrent", "Current", "FirstReleaseDeferred", "Deferred")
        $ChannelXml = Get-ChannelXml

        foreach ($ChannelName in $ChannelList) {
            if ($Channel.ToString().ToLower() -eq $ChannelName.ToLower()) {
                $selectChannel = $ChannelXml.UpdateFiles.baseURL | Where {$_.branch -eq $ChannelName.ToString() }
                $latestVersion = Get-ChannelLatestVersion -ChannelUrl $selectChannel.URL -Channel $ChannelName
                $ChannelShortName = ConvertChannelNameToShortName -ChannelName $ChannelName
                $package = CheckIfPackageExists

                if (!($package)) {
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
                }

                if ($DeploymentPurpose -eq "Default") {
                   $DeploymentPurpose = "Required"  
                }

                $tmpPType = $pType
                if ($CustomName) {
                   $tmpPType += "-$CustomName"
                }

                $Program = Get-CMProgram | Where {$_.Comment.ToLower() -eq $tmpPType.ToLower() }

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

                        $comment = $ProgramType.ToString() + "-" + $ChannelName + "-" + $Bitness.ToString() + "-" + $Collection.ToString()
                        if ($CustomName) {
                           $comment += "-$CustomName" 
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

                                    Update-CMDistributionPoint -PackageId $package.PackageId

                                    Write-Host "`tDeployment created for: $packageName ($ProgramName)"
                                } else {
                                    Write-Host "Could Not find Program in Package for Type: $ProgramType - Channel: $ChannelName" -ForegroundColor White -BackgroundColor Red
                                }
                            } catch {
                                [string]$ErrorMessage = $_.ErrorDetails 
                                if ($ErrorMessage.ToLower().Contains("Could not find property PackageID".ToLower())) {
                                    Write-Host 
                                    Write-Host "Package: $packageName"
                                    Write-Host "The package has not finished deploying to the distribution points." -BackgroundColor Red
                                    Write-Host "Please try this command against once the distribution points have been updated" -BackgroundColor Red
                                } else {
                                    throw
                                }
                            }  
                        } else {
                          Write-Host "`tDeployment already exists for: $packageName ($ProgramName)"
                        }
                   } else {
                        Write-Host "Could Not find Program in Package for Type: $ProgramType - Channel: $ChannelName - Bitness: $Bitness" -ForegroundColor White -BackgroundColor Red
                   }
                } else {
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

)
Begin{
    $currentExecutionPolicy = Get-ExecutionPolicy
	Set-ExecutionPolicy Unrestricted -Scope Process -Force  
    $startLocation = Get-Location
}

Process{
    $SiteCode = GetLocalSiteCode -SiteCode $SiteCode
    $Package = CheckIfPackageExists
    $PkgID = $Package.PackageID

    $query = Get-WmiObject –NameSpace Root\SMS\Site_$SiteCode –Class SMS_DistributionDPStatus –Filter "PackageID='$PkgID'" | Select Name, MessageID, MessageState, LastUpdateDate

    If ($query -eq $null)
    {  
        throw "PackageID not found"
    }

    Foreach ($objItem in $query)

    {
        $DPName = $objItem.Name
        $UpdDate = [System.Management.ManagementDateTimeconverter]::ToDateTime($objItem.LastUpdateDate)

        switch ($objItem.MessageState)
          {
          1 {$Status = "Success"}
          2 {$Status = "In Progress"}
          3 {$Status = "Failed"}
          }

        switch ($objItem.MessageID)
            {
            2303      {$Message = "Content was successfully refreshed"}
            2323      {$Message = "Failed to initialize NAL"}
            2324      {$Message = "Failed to access or create the content share"}
            2330      {$Message = "Content was distributed to distribution point"}
            2354      {$Message = "Failed to validate content status file"}
            2357      {$Message = "Content transfer manager was instructed to send content to Distribution Point"}
            2360      {$Message = "Status message 2360 unknown"}
            2370      {$Message = "Failed to install distribution point"}
            2371      {$Message = "Waiting for prestaged content"}
            2372      {$Message = "Waiting for content"}
            2380      {$Message = "Content evaluation has started"}
            2381      {$Message = "An evaluation task is running. Content was added to Queue"}
            2382      {$Message = "Content hash is invalid"}
            2383      {$Message = "Failed to validate content hash"}
            2384      {$Message = "Content hash has been successfully verified"}
            2391      {$Message = "Failed to connect to remote distribution point"}
            2398      {$Message = "Content Status not found"}
            8203      {$Message = "Failed to update package"}
            8204      {$Message = "Content is being distributed to the distribution Point"}
            8211      {$Message = "Failed to update package"}
            }
        Write-Host "Package $PkgID on $DPName is in '$Status' state"
        Write-Host "Detail: $Message"
        Write-Host "Last Update Date: $UpdDate"
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
		[String]$SourcePath = $NULL
	) 
    Process {
	  $doc = [Xml] (Get-Content $Path)

      $addNode = $doc.Configuration.Add

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

        $CMModulePath = GetCMPSModulePath -CMPSModulePath $CMPSModulePath 
    
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
		[Bool]$UpdateOnlyChangedBits = $true
	) 

    $package = Get-CMPackage | Where { $_.Name -eq $Name }
    if($package -eq $null -or !$package)
    {
        Write-Host "`tCreating Package: $Name"
        $package = New-CMPackage -Name $Name -Path $path -Version $Version
    } else {
        Write-Host "`t`tPackage Already Exists: $Name"        
    }
		
    Write-Host "`t`tSetting Package Properties"

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
		[String]$Version
	) 
    
    if ($Version) {
        $packages = Get-CMPackage | Where { $_.Name -eq $Name -and $_.Version -ne $Version }
        foreach ($package in $packages) {
           $packageName = $package.Name
           $pkversion = $package.Version

           Write-Host "Removing previous version: $packageName - $pkversion"
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
		[String[]] $RequiredPlatformNames = @()

	) 

    $program = Get-CMProgram | Where { $_.PackageID -eq $PackageID -and $_.Comment -eq $Comment }

    if($program -eq $null -or !$program)
    {
        Write-Host "`t`tCreating Program: $Name ..."	        
	    $program = New-CMProgram -PackageId $PackageID -StandardProgramName $Name -DriveMode RenameWithUnc `
                                 -CommandLine $CommandLine -ProgramRunType OnlyWhenUserIsLoggedOn `
                                 -RunMode RunWithAdministrativeRights -UserInteraction $true -RunType Normal 
    } else {
        Write-Host "`t`tProgram Already Exists: $Name"
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

    IF (!(TEST-PATH $Path)) { 
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

function CreateOfficeUpdateShare() {
    [CmdletBinding()]	
    Param
	(
		[Parameter()]
		[String]$Name = "OfficeDeployment$",
		
		[Parameter()]
		[String]$Path = "$env:SystemDrive\OfficeDeployment"
	) 

    IF (!(TEST-PATH $Path)) { 
      $addFolder = New-Item $Path -type Directory 
    }
    
    $ACL = Get-ACL $Path

    $identity = New-Object System.Security.Principal.NTAccount  -argumentlist ("$env:UserDomain\$env:UserName") 
    $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule -argumentlist ($identity,"FullControl","ContainerInherit, ObjectInherit","None","Allow")

    $addAcl = $ACL.AddAccessRule($accessRule)

    $identity = New-Object System.Security.Principal.NTAccount -argumentlist ("$env:UserDomain\Domain Admins") 
    $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule -argumentlist ($identity,"FullControl","ContainerInherit, ObjectInherit","None","Allow")
    $addAcl = $ACL.AddAccessRule($accessRule)

    $identity = "Everyone"
    $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule -argumentlist ($identity,"Read","ContainerInherit, ObjectInherit","None","Allow")
    $addAcl = $ACL.AddAccessRule($accessRule)

    Set-ACL -Path $Path -ACLObject $ACL
    
    $share = Get-WmiObject -Class Win32_share | Where {$_.name -eq "$Name"}
    if (!$share) {
       Create-FileShare -Name $Name -Path $Path
    }

    $sharePath = "\\$env:COMPUTERNAME\$Name"
    return $sharePath
}

function GetSupportedPlatforms([String[]] $requiredPlatformNames){
    $computerName = $env:COMPUTERNAME
    #$assignedSite = $([WmiClass]"\\$computerName\ROOT\ccm:SMS_Client").getassignedsite()
    $siteCode = Get-Site  
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
		[String]$CMPSModulePath = $NULL
	)

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
       throw "Cannot find the ConfigurationManager.psd1 file. Please use the -CMPSModulePath parameter to specify the location of the PowerShell Module"
    }

    return $CMModulePath
}

function Get-Site([string[]]$computerName = $env:COMPUTERNAME) {
    Get-WmiObject -ComputerName $ComputerName -Namespace "root\SMS" -Class "SMS_ProviderLocation" | foreach-object{ 
        if ($_.ProviderForLocalSite -eq $true){$SiteCode=$_.sitecode} 
    } 
    if ($SiteCode -eq "") { 
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
       $scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
     }

     return $scriptPath
 }
}

$scriptPath = GetScriptRoot

$shareFunctionsPath = "$scriptPath\SharedFunctions.ps1"
if ($scriptPath.StartsWith("\\")) {
} else {
    if (!(Test-Path -Path $shareFunctionsPath)) {
        throw "Missing Dependency File SharedFunctions.ps1"    
    }
}
. $shareFunctionsPath
