Add-Type  -ErrorAction SilentlyContinue -TypeDefinition @"
   public enum OfficeCTRVersion
   {
      Office2013,
      Office2016
   }
"@

function Download-OfficeUpdates {
<#
.SYNOPSIS
This method will download the Office updates to a share

.DESCRIPTION
This method is used to download the updates for Office Click-to-Run to a netork share.  This network share could then either be used as a update source for Office Click-to-Run or it could be used as a package source for SCCM.

.PARAMETER Path
The path to the UNC share to download the Office updates to.

.PARAMETER Version
The version of Office 2013 you wish to update to. E.g. 15.0.4737.1003

.PARAMETER Bitness
Specifies if the target installation is 32 bit or 64 bit. Defaults to 64 bit.

.PARAMETER ProductVersion
Specifies Office 2013 vs Office 2016, defaults to 2013, type in either 2013 or 2016 to specify.

.Example
Download-OfficeUpdates 
Default without parameters specified this will create a local folder named 'OfficeUpdates' on the system drive and then create a hidden share named 'OfficeUpdates$'. It will then download the latest Office update to that folder.

.Example
Download-OfficeUpdates -Path "\\Server\OfficeShare"

If you do not want to host the update files on the local server you can specify a UNC share path. The script must be run with a user account that has Read/Write permissions to the share.
 
.Example
Download-OfficeUpdates -Path "\\Server\OfficeShare" -Version "15.0.4737.1003" 
If you specify a Version then the script will download that version.  You can see the version history at https://support.microsoft.com/en-us/gp/office-2013-365-update
#>
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param
    (
        [Parameter(Mandatory=$True)]
	    [OfficeCTRVersion]$ProductVersion = 'Office2013',

	    [Parameter()]
	    [String]$Path = $NULL,

	    [Parameter()]
	    [String]$Version = $NULL,

	    [Parameter()]
	    [String]$Bitness = 'All'
    )
    Begin
    {
        $startLocation = Get-Location
    }
    Process
    {
        if (!$Path) {
           $Path = CreateOfficeUpdateShare
        }

        [String]$UpdateSourceConfigFileName32 = 'Configuration_UpdateSource32.xml'
        [String]$UpdateSourceConfigFileName64 = 'Configuration_UpdateSource64.xml'

        CreateDownloadXmlFile -Path $path -ConfigFileName $UpdateSourceConfigFileName32 -Bitness 32 -Version $version
        CreateDownloadXmlFile -Path $path -ConfigFileName $UpdateSourceConfigFileName64 -Bitness 64 -Version $version

        
        
        if($ProductVersion -eq 'Office2016'){
            $c2rFileName = "Office2016Setup.exe"
        }
        else{
            $c2rFileName = "Office2013Setup.exe"
        }


        Set-Location $PSScriptRoot

        if (!(Test-Path -Path "$Path\$c2rFileName")) {
              Copy-Item -Path ".\$c2rFileName" -Destination $Path
        }
        if (!(Test-Path -Path "$Path\SCO365PPTrigger.exe")) {
              Copy-Item -Path ".\SCO365PPTrigger.exe" -Destination $Path
        }

	    #Connect PowerShell to Share location	
	    Set-Location $path

        Write-Host "Staging the Office ProPlus Update to: $path"
        <# write log#>
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Staging the Office ProPlus Update to: $path"
        Write-Host
         
	    if (($bitness.ToLower() -eq "all") -or ($bitness -eq "32")) {
	        $app = "$path\$c2rFileName" 
	        $arguments = "/download", "$UpdateSourceConfigFileName32"
 
            Write-Host "`tStarting Download of Office Update 32-Bit..." -NoNewline
            <# write log#>
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Starting Download of Office Update 32-Bit..."

	        #run the executable, this will trigger the download of bits to \\ShareName\Office\Data\
	        & $app @arguments

            Write-Host "`tComplete"
            <# write log#>
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Complete"
        }

	    if (($bitness.ToLower() -eq "all") -or ($bitness -eq "64")) {
	        $app = "$path\$c2rFileName" 
	        $arguments = "/download", "$UpdateSourceConfigFileName64"

            Write-Host "`tStarting Download of Office Update 64-Bit..."  -NoNewline
            <# write log#>
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Starting Download of Office Update 64-Bit..."

	        #run the executable, this will trigger the download of bits to \\ShareName\Office\Data\
	        & $app @arguments

            Write-Host "`tComplete"
            <# write log#>
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Complete"
        }

        Write-Host
        Write-Host "The Office Update download has finished"
        <# write log#>
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "The Office Update download has finished"
    }

}

function Setup-SCCMOfficeUpdates {
<#
.SYNOPSIS
Automates the configuration of System Center Configuration Manager (SCCM) to configure Office Click-To-Run Updates

.DESCRIPTION

.PARAMETER Collection
Required. The name of the collection to deploy the updates to.

.PARAMETER $ProductVersion
Required. The version of Office to update.

.PARAMETER path
The UNC Path where the downloaded bits will be stored for updating the target machines.

.PARAMETER version
The version of Office 2013 you wish to update to. E.g. 15.0.4737.1003

.PARAMETER $SiteCode
The 3 Letter Site ID.

.PARAMETER bitness
Specifies if the target installation is 32 bit or 64 bit. Defaults to 64 bit.

.PARAMETER SCCMPSModulePath
Allows the user to specify that full path to the ConfigurationManager.psd1 PowerShell Module. This is especially useful if SCCM is installed in a non standard path.

.Example
.\SetupOfficeUpdatesSCCM.ps1 -version "15.0.4737.1003" -path "\\OfficeShare" -SiteCode "ABC"
Default update Office 2013 to version 15.0.4737.1003

.Example
.\SetupOfficeUpdatesSCCM.ps1 -version "15.0.4737.1003" -path "\\OfficeShare" -bitness "32" -SiteCode "ABC" 
Update Office 2013 to version 15.0.4737.1003 for 32 bit clients
#>

[CmdletBinding(SupportsShouldProcess=$true)]
Param
(
	[Parameter(Mandatory=$True)]
	[String]$Collection,

    [Parameter(Mandatory=$True)]
    [OfficeCTRVersion]$ProductVersion = 'Office2013',

	[Parameter()]
	[String]$Path = $null,

	[Parameter()]
	[String]$Version,

	[Parameter()]
	[String]$SiteCode = $null,
	
	[Parameter()]
	[String]$PackageName = "Office ProPlus Update",
		
	[Parameter()]
	[String]$ProgramName = "Office ProPlus Update",

	[Parameter()]	
	[Bool]$UpdateOnlyChangedBits = $true,

	[Parameter()]
	[String[]] $RequiredPlatformNames = @("All x86 Windows 7 Client", "All x86 Windows 8 Client", "All x86 Windows 8.1 Client", "All Windows 10 Professional/Enterprise and higher (32-bit) Client","All x64 Windows 7 Client", "All x64 Windows 8 Client", "All x64 Windows 8.1 Client", "All Windows 10 Professional/Enterprise and higher (64-bit) Client"),
	
	[Parameter()]
	[string]$DistributionPointGroupName,

	[Parameter()]
	[uint16]$DeploymentExpiryDurationInDays = 15,

	[Parameter()]
	[String]$SCCMPSModulePath = $NULL

)
Begin
{
    $currentExecutionPolicy = Get-ExecutionPolicy
	Set-ExecutionPolicy Unrestricted -Scope Process -Force  
    $startLocation = Get-Location
}
Process
{
    Write-Host
    Write-Host 'Configuring System Center Configuration Manager to Deploy Office ProPlus Updates' -BackgroundColor DarkBlue
    Write-Host
    <# write log#>
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError 'Configuring System Center Configuration Manager to Deploy Office ProPlus Updates'

    if (!$Path) {
         $Path = CreateOfficeUpdateShare
    }

    Set-Location $PSScriptRoot

    if($ProductVersion -eq 'Office2016'){
        $c2rFileName = "Office2016Setup.exe"	    
    }
    else{      
        $c2rFileName = "Office2013Setup.exe"
	}
    $setupExePath = "$path\$c2rFileName"

	Set-Location $startLocation
    Set-Location $PSScriptRoot

    Write-Host "Loading SCCM Module"
    Write-Host ""
    <# write log#>
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Loading SCCM Module"

    #HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\SMS\Setup

    $sccmModulePath = GetSCCMPSModulePath -SCCMPSModulePath $SCCMPSModulePath 
    
    if ($sccmModulePath) {
        Import-Module $sccmModulePath

        if (!$SiteCode) {
           $SiteCode = (Get-ItemProperty -Path "hklm:\SOFTWARE\Microsoft\SMS\Identification" -Name "Site Code").'Site Code'
        }

	    Set-Location "$SiteCode`:"	

        $package = CreateSCCMPackage -Name $PackageName -Path $path -UpdateOnlyChangedBits $UpdateOnlyChangedBits

        CreateSCCMProgram -Name $programName -PackageName $PackageName -Path $path -RequiredPlatformNames $requiredPlatformNames

        Write-Host "Starting Content Distribution"	
        <# write log#>
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Starting Content Distribution"	

        if ($distributionPointGroupName) {
	        Start-CMContentDistribution -PackageName $PackageName -CollectionName $Collection -DistributionPointGroupName $distributionPointGroupName
        } else {
            Start-CMContentDistribution -PackageName $PackageName -CollectionName $Collection
        }

        Write-Host 
        Write-Host "NOTE: In order to deploy the package you must run the function 'Deploy-SCCMOfficeUpdates'." -BackgroundColor Red
        Write-Host "      You should wait until the content has finished distributing to the distribution points." -BackgroundColor Red
        Write-Host "      otherwise the deployments will fail. The clients will continue to fail until the " -BackgroundColor Red
        Write-Host "      content distribution is complete." -BackgroundColor Red

    } else {
    <# write log#>
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Could Not find file ConfigurationManager.psd1"
        throw [System.IO.FileNotFoundException] "Could Not find file ConfigurationManager.psd1"
        
    }
}
End
{
    Set-ExecutionPolicy $currentExecutionPolicy -Scope Process -Force
    Set-Location $startLocation    
}
}

function Deploy-SCCMOfficeUpdates {
<#
.SYNOPSIS
Automates the configuration of System Center Configuration Manager (SCCM) to configure Office Click-To-Run Updates

.DESCRIPTION

.PARAMETER Collection
The target SCCM Collection

.PARAMETER PackageName
The Name of the SCCM package create by the Setup-SCCMOfficeUpdates function

.PARAMETER ProgramName
The Name of the SCCM program create by the Setup-SCCMOfficeUpdates function

.PARAMETER UpdateOnlyChangedBits
Determines whether or not the EnableBinaryDeltaReplication enabled or not

.PARAMETER SCCMPSModulePath
Allows the user to specify that full path to the ConfigurationManager.psd1 PowerShell Module. This is especially useful if SCCM is installed in a non standard path.

.Example
Deploy-SCCMOfficeUpdates -Collection "CollectionName"
Deploys the Package created by the Setup-SCCMOfficeUpdates function
#>
    [CmdletBinding()]	
    Param
	(
		[Parameter(Mandatory=$true)]
		[String]$Collection = "",

		[Parameter()]
		[String]$PackageName = "Office ProPlus Update",

		[Parameter()]
		[String]$ProgramName = "Office ProPlus Update",

		[Parameter()]	
		[Bool]$UpdateOnlyChangedBits = $true,

		[Parameter()]
		[String]$SCCMPSModulePath = $NULL
	) 
Begin
{

}
Process
{
    $sccmModulePath = GetSCCMPSModulePath -SCCMPSModulePath $SCCMPSModulePath 
    
    if ($sccmModulePath) {
        Import-Module $sccmModulePath

        if (!$SiteCode) {
            $SiteCode = (Get-ItemProperty -Path "hklm:\SOFTWARE\Microsoft\SMS\Identification" -Name "Site Code").'Site Code'
        }

	    Set-Location "$SiteCode`:"	

        $package = Get-CMPackage -Name $packageName

        $packageDeploy = Get-CMDeployment | where {$_.PackageId  -eq $package.PackageId }
        if ($packageDeploy.Count -eq 0) {
            Write-Host "Creating Package Deployment for: $packageName"
            <# write log#>
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Creating Package Deployment for: $packageName"
            $dtNow = [datetime]::Now
            $dtNow = $dtNow.AddDays(-1)
            $start = Get-Date -Year $dtNow.Year -Month $dtNow.Month -Day $dtNow.Day -Hour 12 -Minute 0

            $schedule = New-CMSchedule -Start $start -RecurInterval Days -RecurCount 7

     	    Start-CMPackageDeployment -CollectionName $Collection -PackageName $PackageName -ProgramName $ProgramName -StandardProgram  -DeployPurpose Required `
                                      -FastNetworkOption RunProgramFromDistributionPoint -RerunBehavior AlwaysRerunProgram -ScheduleEvent AsSoonAsPossible `
                                      -SlowNetworkOption DoNotRunProgram -SoftwareInstallation $True -SystemRestart $False -Schedule $schedule

        } else {
            Write-Host "Package Deployment Already Exists for: $packageName"
            <# write log#>
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Package Deployment Already Exists for: $packageName"
        }
    }
}
}

function CreateSCCMPackage() {
    [CmdletBinding()]	
    Param
	(
		[Parameter()]
		[String]$Name = "Office ProPlus Update",
		
		[Parameter(Mandatory=$True)]
		[String]$Path,

		[Parameter()]	
		[Bool]$UpdateOnlyChangedBits = $true
	) 

    Write-Host "`tPackage: $Name"

    $package = Get-CMPackage -Name $Name 

    if($package -eq $null -or !$package)
    {
        Write-Host "`t`tCreating Package: $Name"
        <# write log#>
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Creating Package: $Name"
        $package = New-CMPackage -Name $Name  -Path $path
    } else {
        Write-Host "`t`tAlready Exists"	
        <# write log#>
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Already Exists"	
    }
		
    Write-Host "`t`tSetting Package Properties"
    <# write log#>
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Setting Package Properties"

	Set-CMPackage -Name $packageName -Priority High -EnableBinaryDeltaReplication $UpdateOnlyChangedBits

    Write-Host ""

    $package = Get-CMPackage -Name $Name
    return $package
}

function CreateSCCMProgram() {
    [CmdletBinding()]	
    Param
	(
		[Parameter()]
		[String]$PackageName = "Office ProPlus Update",
		
		[Parameter(Mandatory=$True)]
		[String]$Path, 

		[Parameter()]
		[String]$Name = "Office ProPlus Update",
		
		[Parameter()]
		[String[]] $RequiredPlatformNames = @()

	) 

    $program = Get-CMProgram -PackageName $PackageName -ProgramName $Name

    $commandLine = "SCO365PPTrigger.exe -EnableLogging true -C2RArgs `"updatepromptuser=false forceappshutdown=false displaylevel=false`""

    Write-Host "`tProgram: $Name"
    <# write log#>
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Program: $Name"

    if($program -eq $null -or !$program)
    {
        Write-Host "`t`tCreating Program..."	 
        <# write log#>
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Creating Program..."	 
	    $program = New-CMProgram -PackageName $PackageName -StandardProgramName $Name -CommandLine $commandLine -ProgramRunType WhetherOrNotUserIsLoggedOn -RunMode RunWithAdministrativeRights -UserInteraction $false -RunType Hidden
    } else {
        Write-Host "`t`tAlready Exists"
        <# write log#>
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Already Exists"
    }
	
    Write-Host "`t`tSetting Program Properties"
    <# write log#>
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Setting Program Properties"

    $program.CommandLine = $commandLine    
	$program.SupportedOperatingSystems = GetSupportedPlatforms -requiredPlatformNames $requiredPlatformNames

	# Set to use specified client platforms, See - https://msdn.microsoft.com/en-us/library/hh949572.aspx, ProgramFlags
	$anyPlatform = 0x08000000
	$newFlags = $program.ProgramFlags -band (-bnot $anyPlatform)
 
	$program.ProgramFlags = $newFlags
	$program.Put()

    Write-Host ""
}

function CreateOfficeUpdateShare() {
    [CmdletBinding()]	
    Param
	(
		[Parameter()]
		[String]$Name = "OfficeUpdates$",
		
		[Parameter()]
		[String]$Path = "$env:SystemDrive\OfficeUpdates"
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

function CreateDownloadXmlFile([string]$Path, [string]$ConfigFileName, [string]$Bitness, [string]$Version){
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
     $switchVar = ""
     switch ($($R.ReturnValue))
     {
          
          0 { break}
          2 {$switchVar= "Share:$name Path:$path Result:Access Denied"}
          8 {$switchVar= "Share:$name Path:$path Result:Unknown Failure"}
          9 {$switchVar= "Share:$name Path:$path Result:Invalid Name"}
          10 {$switchVar= "Share:$name Path:$path Result:Invalid Level"}
          21 {$switchVar= "Share:$name Path:$path Result:Invalid Parameter"}
          22 {$switchVar= "Share:$name Path:$path Result:Duplicate Share"}
          23 {$switchVar= "Share:$name Path:$path Result:Reedirected Path"}
          24 {$switchVar= "Share:$name Path:$path Result:Unknown Device or Directory"}
          25 {$switchVar= "Share:$name Path:$path Result:Network Name Not Found"}
          default {$switchVar= "Share:$name Path:$path Result:*** Unknown Error ***"}
     }
     <# write log#>
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError $switchVar
}

function GetSCCMPSModulePath() {
    [CmdletBinding()]	
    Param
	(
		[Parameter()]
		[String]$SCCMPSModulePath = $NULL
	)

    [bool]$pathExists = $false

    if ($SCCMPSModulePath) {
       if ($SCCMPSModulePath.ToLower().EndsWith(".psd1")) {
         $sccmModulePath = $SCCMPSModulePath
         $pathExists = Test-Path -Path $sccmModulePath
       }
    }

    if (!$pathExists) {
        $uiInstallDir = (Get-ItemProperty -Path "hklm:\SOFTWARE\Microsoft\SMS\Setup" -Name "UI Installation Directory").'UI Installation Directory'
        $sccmModulePath = Join-Path $uiInstallDir "bin\ConfigurationManager.psd1"

        $pathExists = Test-Path -Path $sccmModulePath
        if (!$pathExists) {
            $sccmModulePath = "$env:ProgramFiles\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1"
            $pathExists = Test-Path -Path $sccmModulePath
        }
    }

    if (!$pathExists) {
       $uiAdminPath = ${env:SMS_ADMIN_UI_PATH}
       if ($uiAdminPath.ToLower().EndsWith("\bin")) {
           $dirInfo = $uiAdminPath
       } else {
           $dirInfo = ([System.IO.DirectoryInfo]$uiAdminPath).Parent.FullName
       }
      
       $sccmModulePath = $dirInfo + "\ConfigurationManager.psd1"
       $pathExists = Test-Path -Path $sccmModulePath
    }

    if (!$pathExists) {
       $sccmModulePath = "${env:ProgramFiles(x86)}\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1"
       $pathExists = Test-Path -Path $sccmModulePath
    }

    if (!$pathExists) {
       $sccmModulePath = "${env:ProgramFiles(x86)}\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1"
       $pathExists = Test-Path -Path $sccmModulePath
    }

    if (!$pathExists) {
    <# write log#>
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Cannot find the ConfigurationManager.psd1 file. Please use the -SCCMPSModulePath parameter to specify the location of the PowerShell Module"
       throw "Cannot find the ConfigurationManager.psd1 file. Please use the -SCCMPSModulePath parameter to specify the location of the PowerShell Module"
    }

    return $sccmModulePath
}

# Specify one of SCCM servers and Site code is returned automatically 
Function Get-Site([string[]]$computerName = $env:COMPUTERNAME) {
    Get-WmiObject -ComputerName $ComputerName -Namespace "root\SMS" -Class "SMS_ProviderLocation" | foreach-object{ 
        if ($_.ProviderForLocalSite -eq $true){$SiteCode=$_.sitecode} 
    } 
    if ($SiteCode -eq "") { 
    <# write log#>
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Sitecode of ConfigMgr Site at " + $ComputerName + " could not be determined."
        throw ("Sitecode of ConfigMgr Site at " + $ComputerName + " could not be determined.") 
    } else { 
        Return $SiteCode 
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