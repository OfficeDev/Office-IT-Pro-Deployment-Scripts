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
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param
    (
        [Parameter()]
        [OfficeChannel[]] $Channels = @(1,2,3),

        [Parameter()]
	    [String]$OfficeFilesPath = $NULL,

        [Parameter()]
        [ValidateSet("en-us","ar-sa","bg-bg","zh-cn","zh-tw","hr-hr","cs-cz","da-dk","nl-nl","et-ee","fi-fi","fr-fr","de-de","el-gr","he-il","hi-in","hu-hu","id-id","it-it",
                    "ja-jp","kk-kh","ko-kr","lv-lv","lt-lt","ms-my","nb-no","pl-pl","pt-br","pt-pt","ro-ro","ru-ru","sr-latn-rs","sk-sk","sl-si","es-es","sv-se","th-th",
                    "tr-tr","uk-ua")]
        [string[]] $Languages = ("en-us"),

        [Parameter()]
        [Bitness] $Bitness = 0
        
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

            Download-OfficeProPlusChannels -TargetDirectory $OfficeFilesPath  -Channels $Channel -Version $latestVersion -UseChannelFolderShortName $true -Languages $Languages -Bitness $Bitness

            $cabFilePath = "$env:TEMP/ofl.cab"
            Copy-Item -Path $cabFilePath -Destination "$OfficeFilesPath\ofl.cab" -Force

            $latestVersion = Get-ChannelLatestVersion -ChannelUrl $selectChannel.URL -Channel $Channel -FolderPath $OfficeFilesPath -OverWrite $true 
         }
       }
    }
}
 
function Create-CMOfficePackage {
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

           Write-Host "Updating Package: $packageName"

           $Path = $existingPackage.PkgSourcePath

           $packageName = "Office 365 ProPlus"
           $ChannelPath = "$Path\$Channel"
           $LocalPath = $existingShare.Path
           $LocalChannelPath = $existingShare.Path + "\SourceFiles"

           [System.IO.Directory]::CreateDirectory($LocalChannelPath) | Out-Null
                          
           if ($OfficeSourceFilesPath) {
                Write-Host "`tUpdating Source Files..."

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
             Write-Host "`tUpdating Deployment Files..."
             Copy-Item -Path $DeploymentFilePath -Destination "$LocalPath" -Force -Recurse
           } else {
             throw "Deployment folder missing: $DeploymentFilePath"
           }

           LoadCMPrereqs -SiteCode $SiteCode -CMPSModulePath $CMPSModulePath

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
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param
    (
        [Parameter()]
        [OfficeChannel[]] $Channels = @(1,2,3),

        [Parameter()]
	    [Bitness]$Bitness = "v32",

	    [Parameter()]
	    [CMDeploymentType]$DeploymentType = "DeployWithScript",

	    [Parameter()]
	    [String]$ScriptName = "CM-OfficeDeploymentScript.ps1",

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

         $platforms = @()

         if ($Bitness -eq "Both") {
            $platforms += "32"
            $platforms += "64"
         } else {
            $platforms += $Bitness.ToString().Replace("v", "")
         }

         foreach ($channel in $Channels) {
             LoadCMPrereqs -SiteCode $SiteCode -CMPSModulePath $CMPSModulePath

             $existingPackage = CheckIfPackageExists
             if (!($existingPackage)) {
                throw "You must run the Create-CMOfficePackage function before running this function"
             }

             foreach ($platform in $platforms) {
                 [string]$CommandLine = ""
                 [string]$ProgramName = ""

                 if ($DeploymentType -eq "DeployWithScript") {
                     $ProgramName = "Deploy $channel Channel With Script - $platform-Bit"
                     $CommandLine = "%windir%\Sysnative\windowsPowershell\V1.0\powershell.exe -ExecutionPolicy Bypass -NoLogo -NonInteractive " + `
                                    "-NoProfile -WindowStyle Hidden -Command .\CM-OfficeDeploymentScript.ps1 -Channel $channel -SourceFileFolder SourceFiles -Bitness $platform"

                 } elseif ($DeploymentType -eq "DeployWithConfigurationFile") {
                     $ProgramName = "Deploy $channel Channel With Configuration File - $platform-Bit"
                     $CommandLine = "Office2016Setup.exe /configure Configuration_UpdateSource.xml"

                 }

                 [string]$packageId = $null

                 $packageId = $existingPackage.PackageId
                 if ($packageId) {
                    $comment = $DeploymentType.ToString() + "-" + $channel + "-" + $platform

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

function Create-CMOfficeChannelChangeProgram {
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
             $CommandLine = "%windir%\Sysnative\windowsPowershell\V1.0\Powershell.exe -ExecutionPolicy Bypass -NoLogo -NonInteractive -NoProfile -WindowStyle Hidden -File .\Change-OfficeChannel.ps1 -Channel $Channel"

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
         $CommandLine = "%windir%\Sysnative\windowsPowershell\V1.0\Powershell.exe -ExecutionPolicy Bypass -NoLogo -NonInteractive -NoProfile -WindowStyle Hidden -File .\Change-OfficeChannel.ps1 -Rollback"

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

         [string]$ProgramName = "Update Office 365 With ConfigMgr"
         [string]$CommandLine = "%windir%\Sysnative\windowsPowershell\V1.0\Powershell.exe -ExecutionPolicy Bypass -NoLogo -NonInteractive -NoProfile -WindowStyle Hidden -File .\Update-Office365Anywhere.ps1"

         $CommandLine += " -WaitForUpdateToFinish " + (Convert-Bool -value $WaitForUpdateToFinish) + ` 
                         " -EnableUpdateAnywhere " + (Convert-Bool -value $EnableUpdateAnywhere) + ` 
                         " -ForceAppShutdown " + (Convert-Bool -value $ForceAppShutdown) + ` 
                         " -UpdatePromptUser " + (Convert-Bool -value $UpdatePromptUser) + ` 
                         " -DisplayLevel " + (Convert-Bool -value $DisplayLevel)

         if ($UpdateToVersion) {
             $CommandLine += "-UpdateToVersion " + $UpdateToVersion
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

         [string]$CommandLine = "%windir%\Sysnative\windowsPowershell\V1.0\Powershell.exe -ExecutionPolicy Bypass -NoLogo -NonInteractive -NoProfile -WindowStyle Hidden -File .\Create-Office365AnywhereTask.ps1"
         [string]$ProgramName = "Update Office 365 With Scheduled Task"

         $CommandLine += " -WaitForUpdateToFinish " + (Convert-Bool -value $WaitForUpdateToFinish) + ` 
                         " -EnableUpdateAnywhere " + (Convert-Bool -value $EnableUpdateAnywhere) + ` 
                         " -ForceAppShutdown " + (Convert-Bool -value $ForceAppShutdown) + ` 
                         " -UpdatePromptUser " + (Convert-Bool -value $UpdatePromptUser) + ` 
                         " -DisplayLevel " + (Convert-Bool -value $DisplayLevel)

         if ($UpdateToVersion) {
             $CommandLine += "-UpdateToVersion " + $UpdateToVersion
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

.PARAMETER path
The UNC Path where the downloaded bits will be stored for installation to the target machines.

.PARAMETER Source
The UNC Path where the downloaded branch bits are stored. Required if source parameter is specified.

.PARAMETER Branch

The update branch to be used with the deployment. Current options are "Business, Current, FirstReleaseBusiness, FirstReleaseCurrent".

.PARAMETER $SiteCode
The 3 Letter Site ID.

.PARAMETER CMPSModulePath
Allows the user to specify that full path to the ConfigurationManager.psd1 PowerShell Module. This is especially useful if CM is installed in a non standard path.

.PARAMETER distributionPoint
Sets which distribution points will be used, and distributes the package.

.Example
Setup-CMOfficeProPlusPackage -Path \\CM-CM\OfficeDeployment -PackageName "Office ProPlus Deployment" -ProgramName "Office2016Setup.exe" -distributionPoint CM-CM.CONTOSO.COM -source \\CM-CM\updates -branch Current
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

        $ChannelList = @("FirstReleaseCurrent", "Current", "FirstReleaseDeferred", "Deferred")
        $ChannelXml = Get-ChannelXml

        foreach ($ChannelName in $ChannelList) {
           if ($Channels -contains $ChannelName) {
               $selectChannel = $ChannelXml.UpdateFiles.baseURL | Where {$_.branch -eq $ChannelName.ToString() }
               $latestVersion = Get-ChannelLatestVersion -ChannelUrl $selectChannel.URL -Channel $ChannelName
               $ChannelShortName = ConvertChannelNameToShortName -ChannelName $ChannelName
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
Automates the configuration of System Center Configuration Manager (CM) to configure Office Click-To-Run Updates

.DESCRIPTION

.PARAMETER Collection
The target CM Collection

.PARAMETER PackageName
The Name of the CM package create by the Setup-CMOfficeProPlusPackage function

.PARAMETER ProgramName
The Name of the CM program create by the Setup-CMOfficeProPlusPackage function

.PARAMETER UpdateOnlyChangedBits
Determines whether or not the EnableBinaryDeltaReplication enabled or not

.PARAMETER CMPSModulePath
Allows the user to specify that full path to the ConfigurationManager.psd1 PowerShell Module. This is especially useful if CM is installed in a non standard path.

.Example
Deploy-CMOfficeProPlusPackage -Collection "CollectionName"
Deploys the Package created by the Setup-CMOfficeProPlusPackage function
#>
    [CmdletBinding()]	
    Param
	(
		[Parameter(Mandatory=$true)]
		[String]$Collection = "",

        [Parameter(Mandatory=$true)]
        [OfficeChannel] $Channel,

        [Parameter()]
	    [BitnessOptions]$Bitness = "v32",

        [Parameter(Mandatory=$true)]
        [CMOfficeProgramType] $ProgramType,
        
	    [Parameter()]
	    [String]$SiteCode = $null,

	    [Parameter()]
	    [String]$CMPSModulePath = $NULL,

    	[Parameter()]
	    [DeploymentPurpose]$DeploymentPurpose = "Default"
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
                    }
                    "RollBack" { 
                         $pType = "RollBack"; 
                         if ($DeploymentPurpose -eq "Default") {
                            $DeploymentPurpose = "Available" 
                         }
                    }
                    "UpdateWithConfigMgr" { 
                         $pType = "UpdateWithConfigMgr"; 
                         if ($DeploymentPurpose -eq "Default") {
                            $DeploymentPurpose = "Required"  }
                         }
                    "UpdateWithTask" { 
                         $pType = "UpdateWithTask"; 
                         if ($DeploymentPurpose -eq "Default") {
                            $DeploymentPurpose = "Required"  
                         }
                    }
                }

                if ($DeploymentPurpose -eq "Default") {
                   $DeploymentPurpose = "Required"  
                }

                $Program = Get-CMProgram | Where {$_.Comment -eq $pType }
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

                        $packageDeploy = get-wmiobject -Namespace "root\sms\site_$SiteCode" -Class SMS_Advertisement  | where {$_.PackageId -eq $package.PackageId -and $_.Comment -eq $comment }
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

                                    Write-Host "Deployment created for: $packageName ($ProgramName)"
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
                          Write-Host "Deployment already exists for: $packageName ($ProgramName)"
                        }
                   } else {
                        Write-Host "Could Not find Program in Package for Type: $ProgramType - Channel: $ChannelName" -ForegroundColor White -BackgroundColor Red
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
