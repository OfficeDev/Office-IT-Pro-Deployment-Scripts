function Create-CMOfficeAddinPackage {
<#

.SYNOPSIS
Automates the configuration of System Center Configuration Manager (CM) to create an Office Click-To-Run Package

.PARAMETER SourceFilesPath
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
	    [String]$SourceFilesPath = $NULL,

        [Parameter()]
	    [bool]$MoveSourceFiles = $false,

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
                               
            if(Test-Path $SourceFilesPath){
                $childItems = Get-ChildItem -Path $SourceFilesPath
                foreach($item in $childItems.Name){
                    $itemPath = Join-Path $SourceFilesPath $item
                    if ($MoveSourceFiles) {   
                        Move-Item -Path $itemPath -Destination $SharePath -Force
                    } else {
                        Copy-Item -Path $itemPath -Destination $SharePath -Recurse -Force
                    }
                }
            } else {
                throw "Source folder missing: $SourceFilesPath"
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

            Write-Host

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
	    [String]$PackageName,      

	    [Parameter()]
	    [String]$SharePath,        

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

            if ($ProgramName.Length -gt 50) {
                throw "CustomName is too long.  Must be less then 50 Characters"
            }

            LoadCMPrereqs -SiteCode $SiteCode -CMPSModulePath $CMPSModulePath

            #$LargeDrv = Get-LargestDrive
            #$LocalPath = $SharePath
            #
            #if($CustomPackageShareName){
            #    $SharePath = "$LargeDrv\$CustomPackageShareName"
            #} else {
            #    $SharePath = "$LargeDrv\OfficeAddinScriptDeployment"
            #}

            $existingPackage = CheckIfPackageExists -PackageName $PackageName
            if (!($existingPackage)) {
               throw "You must run the Create-CMOfficeAddinPackage function before running this function"
            }

            [string]$CommandLine = ""
            [string]$ProgramName = ""
      
            $ProgramName = "Get Office Add-ins"

            $CommandLine = "%windir%\Sysnative\windowsPowershell\V1.0\powershell.exe -ExecutionPolicy Bypass -NoLogo -NonInteractive " + `
                           "-NoProfile -WindowStyle Hidden -Command .\$ScriptName"

            [string]$packageId = $null

            $packageId = $existingPackage.PackageId
            if ($packageId) {
               $comment = "Get a list of Office add-ins and add-in attributes"
            
               CreateCMProgram -Name $ProgramName -PackageID $packageId -CommandLine $CommandLine -Comment $comment -LogFilePath $LogFilePath
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
		[String]$Comment = $null,

        [Parameter()]
        [string]$LogFilePath 

	) 

    $program = Get-CMProgram | ? {$_.PackageID -eq $PackageID -and $_.ProgramName -eq $Name}

    if($program -eq $null -or !$program) {
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

Create-CMOfficeAddinPackage -PackageName $PackageName -SourceFilesPath $SourceFilesPath -MoveSourceFiles $MoveSourceFilesPath -CustomPackageShareName $CustomPackageShareName -SiteCode $SiteCode -CMPSModulePath $CMPSModulePath

Create-CMOfficeAddinProgram -ScriptName $ScriptName -PackageName $PackageName -SharePath $SharePath -SiteCode $SiteCode -CMPSModulePath $CMPSModulePath