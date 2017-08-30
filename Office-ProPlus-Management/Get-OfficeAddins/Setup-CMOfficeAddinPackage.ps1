function Create-CMOfficeAddinPackage {
<#

.SYNOPSIS
    Automates the configuration of System Center Configuration Manager (CM) to create an Office Click-To-Run Package

.PARAMETER PackageName
    The name of the package.

.PARAMETER ScriptFilesPath
    This is the location where the source files are available at

.PARAMETER MoveSourceFiles
    This moves the files from the Source location to the location specified

.PARAMETER CustomPackageShareName
    This sets a custom package share to use

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
        [string]$PackageName = "Get Office Addins",
        
        [Parameter()]
	    [String]$ScriptFilesPath = $NULL,

        [Parameter()]
	    [bool]$MoveScriptFiles = $false,

		[Parameter()]
		[String]$CustomPackageShareName = $null,

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
            Check-AdminAccess

            [bool]$packageCreated = $false

            $existingPackage = CheckIfPackageExists -PackageName $PackageName
            $LargeDrv = Get-LargestDrive

            if($CustomPackageShareName){
                $SharePath = "$LargeDrv\$CustomPackageShareName"
            } else {
                $SharePath = "$LargeDrv\OfficeAddinScriptDeployment"
            }

            $Path = CreateOfficeAddinShare -Path $SharePath
                               
            if(Test-Path $ScriptFilesPath){
                $ScriptFilesFolder = "ScriptFiles"
                $ScriptFilesFolderPath = Join-Path $ScriptFilesPath $ScriptFilesFolder
                $items = Get-ChildItem -Path $ScriptFilesFolderPath
                foreach($item in $items){
                    $itemPath = Join-Path $ScriptFilesFolderPath $item.Name
                    if ($MoveScriptFiles) {   
                        Move-Item -Path $itemPath -Destination $SharePath -Force
                    } else {
                        Copy-Item -Path $itemPath -Destination $SharePath -Recurse -Force
                    }
                }

                Remove-Item -Path $ScriptFilesFolderPath -Force

            } else {
                throw "Source folder missing: $ScriptFilesPath"
            }

            LoadCMPrereqs -SiteCode $SiteCode -CMPSModulePath $CMPSModulePath

            if (!($existingPackage)) {
                $package = CreateCMPackage -Name $PackageName -Path $Path
                $packageCreated = $true
                Write-Host "`tPackage Created: $PackageName"
            } else {
                if(!$packageCreated){
                    Write-Host "`tPackage Already Exists: $packageName"
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

function Create-CMOfficeAddinProgram {
<#
.SYNOPSIS
Automates the configuration of System Center Configuration Manager (CM) to configure Office Click-To-Run Deployment

.DESCRIPTION
Creates a program that can be deployed to clients in a target collection to install Office 365 ProPlus.

.PARAMETER ScriptName
Name the script you would like to use "configuration.xml"

.PARAMETER SiteCode 
The site code you would like to create the package on. If left blank it will default to the current site

.PARAMETER CMPSModulePath
Allows the user to specify that full path to the ConfigurationManager.psd1 PowerShell Module. This is especially useful if CM is installed in a non standard path.

.EXAMPLE
Create-CMOfficeDeploymentProgram -Channels Deferred -DeploymentType DeployWithScript

.EXAMPLE
Create-CMOfficeDeploymentProgram -Channels Current -DeploymentType DeployWithConfigurationFile -ScriptName engineering.xml

#>
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param
    (
	    [Parameter()]
	    [String]$ScriptName = "Get-OfficeAddins.ps1",

	    [Parameter()]
	    [String]$PackageName = "Get Office Addins", 

        [Parameter()]
        [String]$ProgramName = "Get Office Add-ins",    

	    [Parameter()]
	    [String]$SiteCode = $null,

	    [Parameter()]
	    [String]$CMPSModulePath = $NULL,

        [Parameter()]
        [string]$LogFilePath,

        [Parameter(ValueFromPipeLine=$true)]
        [string]$WMIClassName = "Custom_OfficeAddins"
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

            if ($ProgramName.Length -gt 50) {
                throw "CustomName is too long.  Must be less then 50 Characters"
            }

            LoadCMPrereqs -SiteCode $SiteCode -CMPSModulePath $CMPSModulePath

            $existingPackage = CheckIfPackageExists -PackageName $PackageName
            if (!($existingPackage)) {
               throw "You must run the Create-CMOfficeAddinPackage function before running this function"
            }

            [string]$CommandLine = ""
      
            if(!$ProgramName){
                $ProgramName = "Get Office Add-ins"
            }

            $CommandLine = "%windir%\Sysnative\windowsPowershell\V1.0\powershell.exe -ExecutionPolicy Bypass -NoLogo -NonInteractive " + `
                           "-NoProfile -WindowStyle Hidden -Command .\$ScriptName -WMIClassName $WMIClassName"

            [string]$packageId = $null

            $packageId = $existingPackage.PackageId
            if ($packageId) {
               $comment = "Get a list of Office add-ins and add-in attributes"
            
               CreateCMProgram -Name $ProgramName -PackageID $packageId -CommandLine $CommandLine -Comment $comment
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

function Create-CMOfficeAddinTaskProgram {
<#
.SYNOPSIS
    Creates an Office Add-in program.

.DESCRIPTION
    Creates an Office 365 ProPlus program that will create a scheduled task on clients in the target collection.

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
    Create-CMOfficeAddinTaskProgram -StartTime 12:00

    In this exmaple, a program called 'Query Office Add-ins with Task' is created. The program will run on clients in the target collection every Tuesday at 12:00. 

#>
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param
    (
    	[Parameter()]
	    [String]$PackageName,

        [Parameter()]
        [string]$ProgramName,

        [Parameter()]
        [bool]$UseRandomStartTime = $true,

        [Parameter()]
        [string]$RandomTimeStart = "08:00",

        [Parameter()]
        [string]$RandomTimeEnd = "17:00",

        [Parameter()]
        [string]$StartTime,

	    [Parameter()]
	    [String]$SiteCode = $null,

	    [Parameter()]
	    [String]$CMPSModulePath = $NULL
    )
    Begin{
        $currentExecutionPolicy = Get-ExecutionPolicy
	    Set-ExecutionPolicy Unrestricted -Scope Process -Force  
        $startLocation = Get-Location
    }
    Process{
        try{
            Check-AdminAccess

            LoadCMPrereqs -SiteCode $SiteCode -CMPSModulePath $CMPSModulePath

            $existingPackage = CheckIfPackageExists -PackageName $PackageName
            if (!($existingPackage)) {
                throw "You must run the Create-CMOfficePackage function before running this function"
            }

            $CommandLine = "%windir%\Sysnative\windowsPowershell\V1.0\Powershell.exe -ExecutionPolicy Bypass -NoLogo -NonInteractive -NoProfile -WindowStyle Hidden -Command .\Create-QueryOfficeAddinsTask.ps1"
            
            if(!$ProgramName){
                $ProgramName = "Update Office Add-in WMI Class with Task"
            }
                          
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

            $OSSourcePath = "$PSScriptRoot\ScriptFiles\Get-OfficeAddins.ps1"
            $OCScriptPath = "$SharePath\Get-OfficeAddins.ps1"

            $OSSourcePathTask = "$PSScriptRoot\ScriptFiles\Create-QueryOfficeAddinsTask.ps1"
            $OCScriptPathTask = "$SharePath\Create-QueryOfficeAddinsTask.ps1"
 
            if (!(Test-ItemPathUNC -Path $SharePath -FileName "Get-OfficeAddins.ps1")) {
                if(!(Test-Path $OSSourcePath)){
                    throw "Required file missing: $OSSourcePath"
                } else {
                    Copy-ItemUNC -SourcePath $SourcePath -TargetPath $SharePath -FileName "Get-OfficeAddins.ps1"
                }
            }

            if($UseScheduledTask) {
                if (!(Test-ItemPathUNC -Path $SharePath -FileName "Create-QueryOfficeAddinsTask.ps1")) {
                    if(!(Test-Path $OSSourcePathTask)){
                        throw "Required file missing: $OSSourcePathTask"
                    } else {
                        Copy-ItemUNC -SourcePath $OSSourcePathTask  -TargetPath $SharePath -FileName "Create-QueryOfficeAddinsTask.ps1"
                    }
                }
            }

            [string]$packageId = $existingPackage.PackageId
            if($packageId) {
                $comment = "QueryWithTask"

                CreateCMProgram -Name $ProgramName -PackageID $packageId -CommandLine $CommandLine -Comment $comment
            }
        } catch {
            throw;
        }
    }
    End{
        Set-ExecutionPolicy $currentExecutionPolicy -Scope Process -Force
        Set-Location $startLocation    
    }
}

function Distribute-CMOfficeAddinPackage {
<#
.SYNOPSIS
Automates the configuration of System Center Configuration Manager (CM) to configure Office Click-To-Run Updates

.DESCRIPTION
Distributes the Office 365 ProPlus package to the specified Distribution Point or Distribution Point Group.

.PARAMETER DistributionPoint
The distribution point name.

.PARAMETER DistributionPointGroupName
The distribution point group name.

.PARAMETER SiteCode
The 3 Letter Site ID.

.PARAMETER CMPSModulePath
Allows the user to specify that full path to the ConfigurationManager.psd1 PowerShell Module. This is especially useful if CM is installed in a non standard path.

.PARAMETER DistributionPoint
Sets which distribution points will be used, and distributes the package.

.Example
Distribute-CMOfficePackage -DistirbutionPoint cm.contoso.com
Distributes the package 'Office 365 ProPlus' to the distribution point cm.contoso.com

#>
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param
    (
	    [Parameter()]
	    [String]$PackageName = "Get Office Addins",
  
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
            Check-AdminAccess

            $package = CheckIfPackageExists -PackageName $PackageName

            if (!($package)) {
                throw "You must run the Create-CMOfficeAddinPackage function before running this function"
            }

            LoadCMPrereqs -SiteCode $SiteCode -CMPSModulePath $CMPSModulePath

            if ($package) {
                [string]$packageName = $package.Name

                if ($DistributionPointGroupName) {
                    Write-Host "`tStarting Content Distribution for package: $packageName"

	                Start-CMContentDistribution -PackageName $packageName -DistributionPointGroupName $DistributionPointGroupName
                }

                if ($DistributionPoint) {
                    Write-Host "`tStarting Content Distribution for package: $packageName"

                    Start-CMContentDistribution -PackageName $packageName -DistributionPointName $DistributionPoint
                }
            }

            if($WaitForDistributionToFinish){
                Get-CMOfficeDistributionStatus -WaitForDistributionToFinish $true -LogFilePath $LogFilePath
            }
            
            Write-Host 
            Write-Host "NOTE: In order to deploy the package you must run the function 'Deploy-CMOfficeProgram'." -BackgroundColor Red
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

function Deploy-CMOfficeAddinProgram {
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


#>
    [CmdletBinding()]	
    Param
	(
        [Parameter()]
	    [String]$PackageName = "Get Office Addins",

        [Parameter()]
        [String]$ProgramName = "Get Office Add-ins",

		[Parameter(Mandatory=$true)]
		[String]$Collection = "",
    
	    [Parameter()]
	    [String]$SiteCode = $NULL,

	    [Parameter()]
	    [String]$CMPSModulePath = $NULL,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet("Default","Available","Required")] 
        [string[]]$DeploymentPurpose = "Default",

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
            Check-AdminAccess

            $package = CheckIfPackageExists -PackageName $PackageName

            if (!($package)) {
                throw "You must run the Create-CMOfficePackage function before running this function"
            }

            LoadCMPrereqs -SiteCode $SiteCode -CMPSModulePath $CMPSModulePath
            
            $SiteCode = GetLocalSiteCode -SiteCode $SiteCode

            if($DeploymentPurpose -eq "Default"){
                $DeploymentPurpose = "Required"  
            }

            $Program = Get-CMProgram | ? {$_.ProgramName -eq $ProgramName}
            if(!$Program){
                throw "You must run the Create-CMOfficeAddinProgram function before running this function"
            }

            $ProgramName = $Program.ProgramName

            $CollectionName = Get-CMCollection | ? {$_.Name -eq $Collection}
            if(!$CollectionName){
                throw "$Collection is not an available Device Collection to choose from"
            } else {
                $CollectionID = $CollectionName.CollectionID
            }

            if($package){
                if($Program){                        
                    $packageDeploy = Get-WmiObject -Namespace "root\sms\site_$SiteCode" -Class SMS_Advertisement  | ? {$_.PackageId -eq $package.PackageId -and $_.ProgramName -eq $ProgramName -and $_.CollectionID -eq $CollectionID}
                    
                    if($packageDeploy.Count -eq 0){
                        try{
                            $packageId = $package.PackageId
                            $ProgramName = $Program.ProgramName

     	                    Start-CMPackageDeployment -CollectionName $Collection -PackageId $packageId -ProgramName $ProgramName `
                                                      -StandardProgram  -DeployPurpose $DeploymentPurpose -RerunBehavior AlwaysRerunProgram `
                                                      -ScheduleEvent AsSoonAsPossible -FastNetworkOption RunProgramFromDistributionPoint `
                                                      -SlowNetworkOption RunProgramFromDistributionPoint -AllowSharedContent $false -Comment $comment

                            Write-Host "`tDeployment created for: $packageName ($ProgramName)"
    
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
        } catch {
            throw;
        }
    }
    End {
        Set-ExecutionPolicy $currentExecutionPolicy -Scope Process -Force
        Set-Location $startLocation 
    }
}

function CheckIfPackageExists() {
    [CmdletBinding()]	
    Param
	(
        [string]$PackageName
    )
    Begin
    {
        $startLocation = Get-Location
    }
    Process {
       LoadCMPrereqs

       $existingPackage = Get-CMPackage | Where { $_.Name -eq $packageName }
       if ($existingPackage) {
         return $existingPackage
       }

       return $null
    }
}

function Check-AdminAccess() {
    if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`    [Security.Principal.WindowsBuiltInRole] "Administrator"))    {        throw "You do not have Administrator rights to run this script!`nPlease re-run this script as an Administrator!"    }
}

function CreateOfficeAddinShare() {
    [CmdletBinding()]	
    Param
	(
        [Parameter()]
        [String]$Name,

        [Parameter()]
        [String]$Path
	) 

    $Name = ($Path | % {$_.Split("\") | select -Last 1}) -join ' '
    $Name = $Name + '$'
    
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

function Create-FileShare() {
    [CmdletBinding()]	
    Param
	(
		[Parameter()]
		[String]$Name = "",
		
		[Parameter()]
		[String]$Path = ""
	)

    $description = "$name"

    $Method = "Create"
    $sd = ([WMIClass] "Win32_SecurityDescriptor").CreateInstance()

    #AccessMasks:
    #2032127 = Full Control
    #1245631 = Change
    #1179817 = Read

    $userName = "$env:USERDOMAIN\$env:USERNAME"

    #Share with the user
    $ACE = ([WMIClass] "Win32_ACE").CreateInstance()
    $Trustee = ([WMIClass] "Win32_Trustee").CreateInstance()
    $Trustee.Name = $userName
    $Trustee.Domain = $NULL
    #original example assigned this, but I found it worked better if I left it empty
    #$Trustee.SID = ([wmi]"win32_userAccount.Domain='york.edu',Name='$name'").sid   
    $ace.AccessMask = 2032127
    $ace.AceFlags = 3 #Should almost always be three. Really. don't change it.
    $ace.AceType = 0 # 0 = allow, 1 = deny
    $ACE.Trustee = $Trustee 
    $sd.DACL += $ACE.psObject.baseobject 

    #Share with Domain Admins
    $ACE = ([WMIClass] "Win32_ACE").CreateInstance()
    $Trustee = ([WMIClass] "Win32_Trustee").CreateInstance()
    $Trustee.Name = "Domain Admins"
    $Trustee.Domain = $Null
    #$Trustee.SID = ([wmi]"win32_userAccount.Domain='york.edu',Name='$name'").sid   
    $ace.AccessMask = 2032127
    $ace.AceFlags = 3
    $ace.AceType = 0
    $ACE.Trustee = $Trustee 
    $sd.DACL += $ACE.psObject.baseobject    
    
     #Share with the user
    $ACE = ([WMIClass] "Win32_ACE").CreateInstance()
    $Trustee = ([WMIClass] "Win32_Trustee").CreateInstance()
    $Trustee.Name = "Everyone"
    $Trustee.Domain = $Null
    #original example assigned this, but I found it worked better if I left it empty
    #$Trustee.SID = ([wmi]"win32_userAccount.Domain='york.edu',Name='$name'").sid   
    $ace.AccessMask = 1179817 
    $ace.AceFlags = 3 #Should almost always be three. Really. don't change it.
    $ace.AceType = 0 # 0 = allow, 1 = deny
    $ACE.Trustee = $Trustee 
    $sd.DACL += $ACE.psObject.baseobject    

    $mc = [WmiClass]"Win32_Share"
    $InParams = $mc.psbase.GetMethodParameters($Method)
    $InParams.Access = $sd
    $InParams.Description = $description
    $InParams.MaximumAllowed = $Null
    $InParams.Name = $name
    $InParams.Password = $Null
    $InParams.Path = $path
    $InParams.Type = [uint32]0

    $R = $mc.PSBase.InvokeMethod($Method, $InParams, $Null)
    switch ($($R.ReturnValue))
     {
          0 { break}
          2 {Write-Host "Share:$name Path:$path Result:Access Denied" -foregroundcolor red -backgroundcolor yellow;break}
          8 {Write-Host "Share:$name Path:$path Result:Unknown Failure" -foregroundcolor red -backgroundcolor yellow;break}
          9 {Write-Host "Share:$name Path:$path Result:Invalid Name" -foregroundcolor red -backgroundcolor yellow;break}
          10 {Write-Host "Share:$name Path:$path Result:Invalid Level" -foregroundcolor red -backgroundcolor yellow;break}
          21 {Write-Host "Share:$name Path:$path Result:Invalid Parameter" -foregroundcolor red -backgroundcolor yellow;break}
          22 {Write-Host "Share:$name Path:$path Result:Duplicate Share" -foregroundcolor red -backgroundcolor yellow;break}
          23 {Write-Host "Share:$name Path:$path Result:Reedirected Path" -foregroundcolor red -backgroundcolor yellow;break}
          24 {Write-Host "Share:$name Path:$path Result:Unknown Device or Directory" -foregroundcolor red -backgroundcolor yellow;break}
          25 {Write-Host "Share:$name Path:$path Result:Network Name Not Found" -foregroundcolor red -backgroundcolor yellow;break}
          default {Write-Host "Share:$name Path:$path Result:*** Unknown Error ***" -foregroundcolor red -backgroundcolor yellow;break}
     }
}

function CreateCMPackage() {
    [CmdletBinding()]	
    Param
	(
		[Parameter()]
		[String]$Name,
		
		[Parameter(Mandatory=$True)]
		[String]$Path,

        [Parameter()]
        [string]$LogFilePath
	) 

    $package = Get-CMPackage | Where { $_.Name -eq $Name }
    if($package -eq $null -or !$package)
    {
        Write-Host "`tCreating Package: $Name"
        $package = New-CMPackage -Name $Name -Path $path
    } else {
        Write-Host "`t`tPackage Already Exists: $Name"       
    }
		
    Write-Host "`t`tSetting Package Properties"    

	Set-CMPackage -Id $package.PackageId -Priority Normal -CopyToPackageShareOnDistributionPoint $True
    
    $package = Get-CMPackage | ? {$_.Name -eq $Name}

    return $package
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
		[String]$Comment = $null

	) 

    $program = Get-CMProgram | ? {$_.PackageID -eq $PackageID -and $_.ProgramName -eq $Name}

    if($program -eq $null -or !$program) {
        Write-Host "`tCreating Program: $Name ..."	   
     
	    $program = New-CMProgram -PackageId $PackageID -StandardProgramName $Name -DriveMode RenameWithUnc `
                                 -CommandLine $CommandLine -ProgramRunType OnlyWhenUserIsLoggedOn `
                                 -RunMode RunWithAdministrativeRights -UserInteraction $true -RunType Normal
    } else {
        Write-Host "`tProgram Already Exists: $Name"   
    }

    if ($program) {
        Set-CMProgram -InputObject $program -Comment $Comment -StandardProgramName $Name -StandardProgram
    }
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
    $SiteCode = GetLocalSiteCode -SiteCode $SiteCode
    $Package = CheckIfPackageExists -PackageName $PackageName
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
                    $Status
                }
            }else{
                if($trackProgress -notcontains $Status.Status){
                    $trackProgress += $Status.Status
                    $updateRunning = $true
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
                    $Status
                }
            }

            if(($Status.Operation -eq 'Failed') -or`
               ($Status.Operation -eq 'Error')){
                $trackProgress += $Status.Status
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
                    $Status
                }
                else{        
                    if($trackProgress -notcontains $Status.Status){
                        $trackProgress += $Status.Status
                        $updateRunning = $true
                        $Status
                    }
                }
            }

            if($Status.Operation -eq "Failed"){
                $updateRunning = $false
                $Status
            }  
            
        }while($updateRunning -eq $true)
    }else{
        $Status
    }
}

End{
    Set-ExecutionPolicy $currentExecutionPolicy -Scope Process -Force
    Set-Location $startLocation
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
    $query = Get-WmiObject -NameSpace Root\SMS\Site_$SiteCode -Class SMS_DistributionDPStatus -Filter "PackageID='$PkgID'" | Select Name, MessageID, MessageState, LastUpdateDate

    if($query -eq $null){  
        throw "PackageID not found"
    }

    foreach($objItem in $query){

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

function Test-ItemPathUNC() {    [CmdletBinding()]	
    Param
	(	    [Parameter(Mandatory=$true)]
	    [String]$Path,	    [Parameter()]
	    [String]$FileName = $null    )    Process {       $pathExists = $false       if ($FileName) {         $filePath = "$Path\$FileName"         $pathExists = [System.IO.File]::Exists($filePath)       } else {         $pathExists = [System.IO.Directory]::Exists($Path)         if (!($pathExists)) {            $pathExists = [System.IO.File]::Exists($Path)         }       }       return $pathExists;    }}

function Copy-ItemUNC() {    [CmdletBinding()]	
    Param
	(	    [Parameter(Mandatory=$true)]
	    [String]$SourcePath,	    [Parameter(Mandatory=$true)]
	    [String]$TargetPath,	    [Parameter(Mandatory=$true)]
	    [String]$FileName    )    Process {       $drvLetter = FindAvailable       $Network = New-Object -ComObject "Wscript.Network"       try {           if (!($drvLetter.EndsWith(":"))) {               $drvLetter += ":"           }           $target = $drvLetter + "\"           $Network.MapNetworkDrive($drvLetter, $TargetPath)           Copy-Item -Path $SourcePath -Destination $target -Force       } finally {         $Network.RemoveNetworkDrive($drvLetter)       }    }}function FindAvailable() {
   #$drives = Get-PSDrive | select Name
   $drives = Get-WmiObject -Class Win32_LogicalDisk | select DeviceID

   for($n=90;$n -gt 68;$n--) {
      $letter= [char]$n + ":"
      $exists = $drives | where { $_.DeviceID -eq $letter }
      if ($exists) {
        if ($exists.Count -eq 0) {
            return $letter
        }
      } else {
        return $letter
      }
   }
   return $null
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