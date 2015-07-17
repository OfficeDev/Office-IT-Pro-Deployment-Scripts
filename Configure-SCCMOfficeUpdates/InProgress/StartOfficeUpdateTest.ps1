
function GetSupportedPlatforms
{
    $computerName = $env:COMPUTERNAME
    $siteCode = $([WmiClass]"\\$computerName\ROOT\ccm:SMS_Client").getassignedsite().sSiteCode

    $requiredPlatformNames = @("All x86 Windows 7 Client", "All x64 Windows 7 Client", "All x86 Windows 8 Client", "All x64 Windows 8 Client", "All Windows 10 Professional/Enterprise and higher (32-bit) Client", "All Windows 10 Professional/Enterprise and higher (64-bit) Client")

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

New-CMPackage -Name "O365Update" -Path "\\officeautosc-dc\Shares\Office" #Path - Make Parameter 

$package = Get-CMPackage -Name "O365Update"

Set-CMPackage -Name "O365Update" -Priority High -EnableBinaryDeltaReplication $true #EnableBinaryDeltaReplication - Make Parameter

New-CMProgram -PackageName "O365Update" -StandardProgramName "Office Updater" -CommandLine 'SCO365PPTrigger.exe -EnableLogging true -C2RArgs "Setup.exe /Configure confguration.xml"' -ProgramRunType WhetherOrNotUserIsLoggedOn -RunMode RunWithAdministrativeRights -UserInteraction $false -RunType Hidden -WorkingDirectory "\\officeautosc-dc\shares\office"

$program = Get-CMProgram -PackageName "O365Update" -ProgramName "Office Updater"
$program.SupportedOperatingSystems = GetSupportedPlatforms
# Set to use specified client platforms, See - https://msdn.microsoft.com/en-us/library/hh949572.aspx, ProgramFlags
$anyPlatform = 0x08000000 #Define the flag as a Constant since we can't find an enum for it.
$newFlags = $program.ProgramFlags -band (-bnot $anyPlatform) 
$program.ProgramFlags = $newFlags
$program.Put()

Start-CMContentDistribution -PackageName "O365Update" -CollectionName "TestCollection" -DistributionPointGroupName "TestDPGroup" 


Start-CMPackageDeployment -CollectionName "TestCollection" -PackageName "O365Update" -ProgramName "Office Updater" -StandardProgram -DeploymentAvailableDateTime ([datetime]::Now.ToString()) -DeploymentExpireDateTime ([datetime]::Now.AddDays(5)).ToString() -DeployPurpose Required -FastNetworkOption DownloadContentFromDistributionPointAndRunLocally -RerunBehavior RerunIfFailedPreviousAttempt -ScheduleEvent AsSoonAsPossible -SlowNetworkOption DoNotRunProgram -SoftwareInstallation $True -SystemRestart $True
