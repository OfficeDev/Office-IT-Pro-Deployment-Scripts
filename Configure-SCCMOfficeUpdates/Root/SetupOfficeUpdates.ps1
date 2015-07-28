<#
.SYNOPSIS
Sets up SCCM Package Deployment for Office Updates. 
.DESCRIPTION
Given a UNC Path, Device Collection Name, a config file, and a Distribution Point Group Name, this cmdlet creates SCCM Package Deployment to update Target Machines.
.PARAMETER packageName
The name for the SCCM Package, Defaults to 'O365Update'
.PARAMETER path
The UNC Path where the bits are stored for updating the target machines.
.PARAMETER UpdateOnlyChangedBits
Specifies if the only changed bits should be updated. This is useful in low bandwidth / resources scenario. Defaults to True
.PARAMETER programName
The name for the SCCM Program under the package. Defaults to 'Office Updater'.
.PARAMETER configFileName
The config file that is used to update the target machines to the intended version.
.PARAMETER requiredPlatformNames
The platforms to support for the SCCM Program. Defaults to Windows 7, Windows 8, Windows 8.1, and Windows 10 clients, both 32-bit and 64-bit.
.PARAMETER collectionToUse
The name of the device collection in SCCM to target.
.PARAMETER distributionPointGroupName
The name of the Distribution Point Group to use for deployment.
.PARAMETER deploymentExpiryDurationInDays
The duration in days for which the content will be available for the deployment.
.Example
SetupOfficeUpdates -path '\\OfficeShare' -collectionToUse 'TestCollection' -distributionPointGroupName 'TestDPGroup' -configFileName TargetConfig.xml
Default update Office 2013 to version 15.0.4737.1003
.Example
.\SetupOfficeUpdatesSCCM.ps1 -version "15.0.4737.1003" -path "\\OfficeShare" -bitness "32" -siteId "ABC" -UpdateSourceConfigFileName "SourceConfig.xml" -UpdateTestGroupConfigFileName "TargetConfig.xml" 
Update Office 2013 to version 15.0.4737.1003 for 32 bit clients
.Inputs
System.String
System.String
System.Bool
System.String
System.String
System.String
System.String
System.String
System.Uint16
.Notes
Additional explaination. Long and indepth examples should also go here.
.Link
Add link here
#>
function SetupOfficeUpdates
{
    [CmdletBinding()]	
    Param
	(
		[Parameter()]
		[String]$packageName = "O365Update",
		
		[Parameter(Mandatory=$True)]
		[String]$path,

		[Parameter()]	
		[Bool]$UpdateOnlyChangedBits = $true,

		[Parameter()]
		[String]$programName = "Office Updater",
		
		[Parameter(Mandatory=$True)]
		[String]$configFileName, 
		
		[Parameter()]
		[String[]] $requiredPlatformNames = @("All x86 Windows 7 Client", "All x64 Windows 7 Client", "All x86 Windows 8 Client", "All x64 Windows 8 Client", "All x86 Windows 8.1 Client", "All x64 Windows 8.1 Client", "All Windows 10 Professional/Enterprise and higher (32-bit) Client", "All Windows 10 Professional/Enterprise and higher (64-bit) Client"),
		
		[Parameter(Mandatory=$True)]
		[String]$collectionToUse,
		
		[Parameter(Mandatory=$True)]
		[string]$distributionPointGroupName,

		[Parameter()]
		[uint16]$deploymentExpiryDurationInDays = 15
	) 
  
    $package = Get-CMPackage -Name $packageName

    if($package -eq $null -or !$package)
    {
        Write-Host 'Creating Package'
        $package = New-CMPackage -Name $packageName -Path $path
    }
		
    Write-Host 'Setting Package Properties'

	Set-CMPackage -Name $packageName -Priority High -EnableBinaryDeltaReplication $UpdateOnlyChangedBits

    $program = Get-CMProgram -PackageName $packageName -ProgramName $programName

    $commandLine = "SCO365PPTrigger.exe -EnableLogging true -C2RArgs `"Setup.exe /Configure $configFileName`""

    if($program -eq $null -or !$program)
    {
        Write-Host 'Creating Program'	        
	    $program = New-CMProgram -PackageName $packageName -StandardProgramName $programName -CommandLine $commandLine -ProgramRunType WhetherOrNotUserIsLoggedOn -RunMode RunWithAdministrativeRights -UserInteraction $false -RunType Hidden -WorkingDirectory $path
    }
	
    Write-Host 'Setting Program Properties'

    $program.CommandLine = $commandLine    
	$program.SupportedOperatingSystems = GetSupportedPlatforms -requiredPlatformNames $requiredPlatformNames
	# Set to use specified client platforms, See - https://msdn.microsoft.com/en-us/library/hh949572.aspx, ProgramFlags
	$anyPlatform = 0x08000000 #Define the flag as a Constant since we can't find an enum for it.
	$newFlags = $program.ProgramFlags -band (-bnot $anyPlatform) 
	$program.ProgramFlags = $newFlags
	$program.Put()

    Write-Host 'Starting Content Distribution'	

	Start-CMContentDistribution -PackageName $packageName -CollectionName $collectionToUse -DistributionPointGroupName $distributionPointGroupName

    Write-Host 'Starting Deployment'	
	
	Start-CMPackageDeployment -CollectionName $collectionToUse -PackageName $packageName -ProgramName $programName -StandardProgram -DeploymentAvailableDateTime ([datetime]::Now.ToString()) -DeploymentExpireDateTime ([datetime]::Now.AddDays($deploymentExpiryDurationInDays)).ToString() -DeployPurpose Required -FastNetworkOption RunProgramFromDistributionPoint -RerunBehavior RerunIfFailedPreviousAttempt -ScheduleEvent AsSoonAsPossible -SlowNetworkOption DoNotRunProgram -SoftwareInstallation $True -SystemRestart $True
}

function GetSupportedPlatforms([String[]] $requiredPlatformNames)
{
    $computerName = $env:COMPUTERNAME
    $siteCode = $([WmiClass]"\\$computerName\ROOT\ccm:SMS_Client").getassignedsite().sSiteCode  
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


 



