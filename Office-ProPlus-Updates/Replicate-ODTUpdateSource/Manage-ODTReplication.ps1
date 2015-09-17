<#
.SYNOPSIS
Manage the download and replication of the latest
Click-to-Run builds.

.DESCRIPTION
The functions available will provide IT Pros with methods for polling 
the CDN daily for newer C2R builds, replicate the folders from 
the source to specified remote shares on the network, and manage the list
of available shares to replicate to.

.FUNCTIONS
Download-ODTOfficeFiles

Replicate-ODTOfficeFiles

Schedule-ODTRemoteShareReplicationTask

Add-ODTRemoteUpdateSource

Remove-ODTRemoteUpdateSource

List-ODTRemoteUpdateSource

.LINKS
Overview on using the ODT
https://technet.microsoft.com/en-us/library/jj219422.aspx

Download the ODT
http://www.microsoft.com/en-us/download/details.aspx?id=36778

Reference for Click-to-Run configuration.xml files
https://technet.microsoft.com/en-us/library/JJ219426.aspx

Reference for creating scheduled tasks
https://msdn.microsoft.com/en-us/library/windows/desktop/bb736357(v=vs.85).aspx

.NOTES
Before using Download-ODTOfficeFiles verify you have the
correct Setup.exe file to download C2R builds from
the CDN.
#>


$OfficeCTRVersion = @"
   public enum OfficeCTRVersion
   {
      Office2013,
      Office2016
   }
"@ 
Add-Type -TypeDefinition $OfficeCTRVersion

$OfficeCTRVersionSel = @"
   public enum OfficeVersionSelection
   {
      All,
      Office2013,
      Office2016
   }
"@ 
Add-Type -TypeDefinition $OfficeCTRVersionSel

$Schedule = @"
   public enum Schedule
   {
      MONTHLY
   }
"@ 
Add-Type -TypeDefinition $Schedule

$Modifier = @"
   public enum Modifier
   {
      FIRST,SECOND,THIRD,FOURTH,LAST
   }
"@ 
Add-Type -TypeDefinition $Modifier

$Days = @"
   public enum Days
   {
      MON,TUE,WED,THU,FRI,SAT,SUN
   }
"@ 
Add-Type -TypeDefinition $Days

$ReplDir = @"
   public enum ReplicationDirection
   {
      Push,
      Pull
   }
"@ 
Add-Type -TypeDefinition $ReplDir


function Start-ODTDownload() {
<#

.SYNOPSIS
Download the latest C2R builds from the CDN.

.DESCRIPTION
A Configuration.xml file is used to download the latest C2r 
build. The appropriate Setup.exe file, provided by Microsoft, 
will need to be used when starting the download. If no taskname
is specified in the parameter the download will begin. If a 
taskname is specified a scheduled task will be created on the
computer to poll the CDN daily for the latest C2R builds.

.PARAMETER OfficeVersion
The version of Office used for the ODT

.PARAMETER XmlConfigPath
Path to the Configuration xml file located on a shared folder

.PARAMETER TaskName
The name of the task created on the source computer

.EXAMPLE
Download-ODTOfficeFiles -OfficeVersion 2013 -XmlConfigPath "\\Server1\ODT Replication"
The Configuration.xml specified will begin the C2R download.

.EXAMPLE
Download-ODTOfficeFiles -OfficeVersion 2013 -XmlConfigPath "C:\ODT Replication" -TaskName "ODT CDN Poll" -ScheduledTime 03:00
A task will be created on the host machine to download the latest C2R builds daily at 3:00am.

#>
    param(
        [Parameter()]
        [OfficeVersionSelection]$OfficeVersion = "Office2013",

        [Parameter()]
        [string] $UpdateVersion = $null,

        [Parameter()]
        [string] $XmlConfigPath = "$PSScriptRoot\configuration.xml",

        [Parameter()]
        [int]$NumberOfVersionsToKeep = 2
    )  

    Begin {

    }

    Process {
        $officeVersions = @()

        if ($OfficeVersion -eq "All") {
           $officeVersions += "Office2013"
           $officeVersions += "Office2016"
        } else {
           $officeVersions += $OfficeVersion
        }

        foreach ($offVersion in $officeVersions) {
            switch($offVersion){
               Office2013 { $odtExtPath = "$PSScriptRoot\Office2013Setup.exe" }
               Office2016 { $odtExtPath = "$PSScriptRoot\Office2016Setup.exe" }
            }

            $progDirPath = "$env:ProgramFiles\Office Update Replication\$offVersion"
            [system.io.directory]::CreateDirectory($progDirPath) | Out-Null

            Write-Host "Downloading `"$offVersion`" Latest 32-Bit Version..." -NoNewline
            $download32 = "$odtExtPath /download $XmlConfigPath"
            Set-ODTAdd -TargetFilePath "$XmlConfigPath" -Bitness 32 -SourcePath $progDirPath | Out-Null

            if ($UpdateVersion) {
              Set-ODTAdd -TargetFilePath "$XmlConfigPath" -Version $UpdateVersion
            }

            Invoke-Expression $download32
            Write-Host "Completed"

            Write-Host "Downloading `"$offVersion`" Latest 64-Bit Version..." -NoNewline
            $download64 = "$odtExtPath /download $XmlConfigPath"
            Set-ODTAdd -TargetFilePath "$XmlConfigPath" -Bitness 64 -SourcePath $progDirPath | Out-Null

            if ($UpdateVersion) {
              Set-ODTAdd -TargetFilePath "$XmlConfigPath" -Version $UpdateVersion
            }

            Invoke-Expression $download64

            Write-Host "Completed"
            Write-Host

            Start-OfficeUpdateSourceCleanup -NumberOfVersionsToKeep $NumberOfVersionsToKeep -OfficeVersion $offVersion
        }

    }
}

function New-ODTDownloadSchedule() {
    param(
        [OfficeCTRVersion]$OfficeVersion = "Office2013",
        [string] $XmlConfigPath = "$PSScriptRoot\configuration.xml",
        [string] $ScheduledTime32Bit = "19:00",
        [string] $ScheduledTime64Bit = "18:00"
    )  
    
    Process {
        $progDirPath = "$env:ProgramFiles\Office Update Replication"
        [system.io.directory]::CreateDirectory($progDirPath) | Out-Null

        switch($OfficeVersion){
          Office2013 { 
             if (!(Test-Path -Path "$env:ProgramFiles\Office Update Replication\Office2013Setup.exe")) {
                Copy-Item -Path "$PSScriptRoot\Office2013Setup.exe" -Destination "$env:ProgramFiles\Office Update Replication\Office2013Setup.exe" -Force | Out-Null
             }
          }
          Office2016 { 
             if (!(Test-Path -Path "$env:ProgramFiles\Office Update Replication\Office2016Setup.exe")) {
               Copy-Item -Path "$PSScriptRoot\Office2016Setup.exe" -Destination "$env:ProgramFiles\Office Update Replication\Office2016Setup.exe" -Force -ErrorAction SilentlyContinue | Out-Null
             }
          }
        }

        Copy-Item -Path $XmlConfigPath -Destination "$env:ProgramFiles\Office Update Replication\$OfficeVersion\configuration32.xml" -Force | Out-Null
        Copy-Item -Path $XmlConfigPath -Destination "$env:ProgramFiles\Office Update Replication\$OfficeVersion\configuration64.xml" -Force | Out-Null

        Set-ODTAdd -TargetFilePath "$env:ProgramFiles\Office Update Replication\$OfficeVersion\configuration32.xml" -SourcePath $progDirPath -Bitness 32 | Out-Null
        Set-ODTAdd -TargetFilePath "$env:ProgramFiles\Office Update Replication\$OfficeVersion\configuration64.xml" -SourcePath $progDirPath -Bitness 64 | Out-Null

        switch($OfficeVersion){
          Office2013 { 
            $odtCmd32 = "\`"$progDirPath\Office2013Setup.exe\`" /Download \`"$env:ProgramFiles\Office Update Replication\$OfficeVersion\configuration32.xml\`"" 
            $odtCmd64 = "\`"$progDirPath\Office2013Setup.exe\`" /Download \`"$env:ProgramFiles\Office Update Replication\$OfficeVersion\configuration64.xml\`"" 
          }
          Office2016 { 
            $odtCmd32 = "\`"$progDirPath\Office2016Setup.exe\`" /Download \`"$env:ProgramFiles\Office Update Replication\$OfficeVersion\configuration32.xml\`"" 
            $odtCmd64 = "\`"$progDirPath\Office2016Setup.exe\`" /Download \`"$env:ProgramFiles\Office Update Replication\$OfficeVersion\configuration64.xml\`"" 
          }
        }

        [string] $computer = $env:COMPUTERNAME
     
        $TaskName32 = "Microsoft\OfficeC2R\$OfficeVersion ODT Download 32-Bit"
        $TaskName64 = "Microsoft\OfficeC2R\$OfficeVersion ODT Download 64-Bit"

        $scheduledTaskAdd32 = "schtasks /create /ru System /tn '$TaskName32' /tr '$odtCmd32' /sc Monthly /mo SECOND /D TUE /st $ScheduledTime32Bit /f"
        $scheduledTaskDel32 = "schtasks /delete /tn '$TaskName32' /f"

        $scheduledTaskAdd64 = "schtasks /create /ru System /tn '$TaskName64' /tr '$odtCmd64' /sc Monthly /mo SECOND /D TUE /st $ScheduledTime64Bit /f"
        $scheduledTaskDel64 = "schtasks /delete /tn '$TaskName64' /f"

        if (findScheduledTask -OfficeVersion $OfficeVersion -Bitness 32) {
           Invoke-Expression $scheduledTaskDel32
        }

        if (findScheduledTask -OfficeVersion $OfficeVersion -Bitness 64) {
           Invoke-Expression $scheduledTaskDel64
        }

        Invoke-Expression $scheduledTaskAdd32
        Invoke-Expression $scheduledTaskAdd64
    }
}

function Remove-ODTDownloadSchedule() {
    param(
        [OfficeCTRVersion]$OfficeVersion = "Office2013"
    )  
    
    Process {
        [string] $computer = $env:COMPUTERNAME
     
        $TaskName32 = "Microsoft\OfficeC2R\$OfficeVersion ODT Download 32-Bit"
        $TaskName64 = "Microsoft\OfficeC2R\$OfficeVersion ODT Download 64-Bit"

        $scheduledTaskDel32 = "schtasks /delete /tn '$TaskName32' /f"
        $scheduledTaskDel64 = "schtasks /delete /tn '$TaskName64' /f"

        if (findScheduledTask -OfficeVersion $OfficeVersion -Bitness 32) {
           Invoke-Expression $scheduledTaskDel32
        } else {
        Write-Host "Task `"$TaskName32`" does not exist"
    }

        if (findScheduledTask -OfficeVersion $OfficeVersion -Bitness 64) {
           Invoke-Expression $scheduledTaskDel64
        } else {
           Write-Host "Task `"$TaskName64`" does not exist"
        }
    }
}


function Start-ODTFileReplication() {
<#

.SYNOPSIS
Replicate the source folder with a list of shared folders on the domain.

.DESCRIPTION
Provide the source and a log file containing the shared folders to replicate
to. A comparison will be made between the source and destination, and if the 
source contains folders not in the destination share the source will be
copied via Robocopy to the destination.

.PARAMETER Source
The source folder hosting the C2R builds.

.PARAMETER ODTShareNameLogFile
The name of the csv file containing a list of shared folders.

.EXAMPLE
Replicate-ODTOfficeFiles -Source "\\Server1\ODT Replication" -ODTShareNameLogFile "\\Server1\ODT Replication\ODTRemoteShares.csv"
The source folder and destination folder(s) will be compared. If the source folder
has updated files or folders they will be copied to each destination.

#>
    Param(
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string[]]$ShareName
    )

    Process {
        $progDirPath = "$env:ProgramFiles\Office Update Replication"
        [system.io.directory]::CreateDirectory($progDirPath) | Out-Null

        $ODTShareNameLogFile = "$progDirPath\ODTReplication.csv"

        $remoteShares = Import-Csv $ODTShareNameLogFile

        foreach($share in $ShareName){
           $existingShares = $remoteShares | where { $_.ShareName.ToLower() -eq $share.ToLower() }
           if ($existingShares.Length -eq 0) {
              throw "Remote Share `"$share`" is not added as a remote source"
           }
        }

        if (!(Test-Path -Path $progDirPath)) {
           throw "Source Path '$Source' does not exist run Start-ODTDownload cmdlet to create it"
        }

        foreach($share in $ShareName){  
            $chkExisting = $remoteShares | where { $_.ShareName.ToLower() -eq $share.ToLower() }
            if ($chkExisting) {
                $existingShare = $chkExisting[0]   
            }

            $officeVersion = $existingShare.OfficeVersion
            $Source = "$progDirPath\$officeVersion\Office"
            
            if (!(Test-Path -Path $Source)) { 
               throw "Before you can replicate the Office ProPlus Click-To-Run Files you must first run the Start-ODTDownload function" 
            }
          
            $sourceFolderPath = $Source

            [system.io.directory]::CreateDirectory($sourceFolderPath) | Out-Null

            $destinationFolder = Get-ChildItem "$share\Office" -Recurse
            $sourceFolder = Get-ChildItem $sourceFolderPath -Recurse

            if($destinationFolder -ne $null){          
                $comparison = Compare-Object -ReferenceObject $sourceFolder -DifferenceObject $destinationFolder -IncludeEqual

                if($comparison.SideIndicator -eq "<="){
                    Write-Host "Copying Office Updates to '$share\Office'"
                    Copy-WithProgress -Source $sourceFolderPath -Destination "$share\Office" 
                    Write-Host
                }
                elseif($comparison.SideIndicator -eq "=="){
                    Write-Host "The remote update source `"$share`" is up to date."
                }
            }
            elseif($destinationFolder -eq $null){
                Write-Host "Copying Office Updates to '$share\Office'"
                Copy-WithProgress -Source $sourceFolderPath -Destination "$share\Office" 
                Write-Host
            }                         
        }
    }
}

function Enable-ODTRemoteUpdateSourceReplication() {
<#
.SYNOPSIS
Create a scheduled task on the remote computer to copy the 
C2R folders from the source on a monthly schedule.

.DESCRIPTION
Given a computer name, source, taskname and the necessary commands
for the task to operate (Schedule,Modifier,Days,StartTime) a scheduled
task can be created on the remote computers to copy the files from
the source.

.PARAMETER RemoteShare
LIst of computers to create the shceduled task on.

.PARAMETER Schedule
A trigger for the script to run Monthly. "MONTHLY" will autopopulate.

.PARAMETER Modifier
The value that refines the scheduled frequency. The list of available
modifiers are FIRST,SECOND,THIRD,FOURTH,LAST.

.PARAMETER Days
Provide the day of week for the task to run on. The list of available
days are MON,TUE,WED,THU,FRI,SAT,SUN.

.PARAMETER StartTime
The time of day the task will run. The hour format is 24-hour (HH:mm)
If no StartTime is given the time will default to the time the task is created.

.EXAMPLE 
Schedule-ODTRemoteShareReplicationTask -ComputerName Computer1,Computer2 -Source "\\Server1\ODT Replication" -TaskName "ODT Replication" -Schedule MONTHLY -Modifier SECOND -Days WED -StartTime 03:00
A task will be created on Computer1 and Computer2 called "ODT Replication"
that will copy the folders from "\\Server1\ODT Replication" every month on
the second Wednesday at 3:00am.


#>
    [CmdletBinding(SupportsShouldProcess=$True)]
    Param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$true)]
        [string[]] $ShareName,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [ReplicationDirection] $ReplicationDirection = "Push",

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [System.Management.Automation.PSCredential] $Credential,

        [Parameter()]
        [Schedule] $Schedule = "MONTHLY",

        [Parameter()]
        [Modifier] $Modifier = "SECOND",

        [Parameter()]
        [Days] $Days = "TUE",

        [Parameter()]
        [string] $StartTime = "20:00", 
         
        [Parameter()]
        [int] $UpdateInterval = 30
    )

    Process {

        $progDirPath = "$env:ProgramFiles\Office Update Replication"
        [system.io.directory]::CreateDirectory($progDirPath) | Out-Null
        $Source = "$progDirPath\Office"
        $ODTShareNameLogFile = "$progDirPath\ODTReplication.csv"

        $remShares = $null
        foreach($remotePath in $ShareName) {
  
            $serverName= $remotePath.Split("\")[2]
            $shareRoot = $remotePath.Split("\")[3]
            $shareName= $remotePath.Replace("\\$serverName\", "").Replace("\", "-")

            $remoteShares = Get-ODTRemoteUpdateSource | Select *
            [datetime]$dtStartTime = getSecondTuesday
            for ($i=0;$i -lt $remoteShares.Length;$i++) {
                $dtStartTime =$dtStartTime.AddMinutes($UpdateInterval)

                $dtHour = $dtStartTime.Hour.ToString()
                $dtMinute = $dtStartTime.Minute.ToString()
                if ($dtMinute.Length -eq 1) { $dtMinute = $dtMinute + "0" }

                [DateTime]$chkStartDate = Get-Date -Hour $dtStartTime.Hour -Minute $dtMinute -Second 0

                $shortTime = $chkStartDate.ToLongTimeString()
                $existingTimes = $remoteShares | Where {$_.StartTime -eq $shortTime }

                $tmpStartTime = $dtHour + ":" + $dtMinute

                if ($existingTimes.Length -eq 0) {
                   $StartTime = $tmpStartTime
                   break
                }
            }

            $TaskName = "Microsoft\OfficeC2R\ODT Replication - $serverName - $shareName"

            $existingTask = findReplScheduledTask -ServerName $serverName -ShareName $shareName
            if ($existingTask) {
                $taskCsv = & schtasks /query /tn $TaskName /v /fo csv
                if ($taskCsv) {
                   $taskCsv   | Out-File -FilePath "$env:temp\TmpSchTask.csv"
	               $importTasks = Import-Csv -Path "$env:temp\TmpSchTask.csv"
        
                   $StartTime = $importTasks.'Start Time' 
                   $StartTime = "{0:HH:mm}" -f [datetime]$StartTime     
                }
            }

            $existingShare = $null
            $chkExisting = $remoteShares | where { $_.ShareName.ToLower() -eq $remotePath.ToLower() }
            if ($chkExisting) {
                $existingShare = $chkExisting[0]   
            }

            $officeVersion = $existingShare.OfficeVersion
            $Source = "$progDirPath\$officeVersion\Office"

            $localComputerName = $env:COMPUTERNAME
            $destRemPath = $remotePath + "\Office"

            try {
                if ($ReplicationDirection -eq "Push") {
                    Grant-SmbShareAccess -name $shareRoot -CimSession $serverName -AccountName "$localComputerName$" -AccessRight Full –Force -ErrorAction Stop | Out-Null
                    Grant-SmbShareAccess -name $shareRoot -CimSession $serverName -AccountName "Administrators" -AccessRight Full –Force -ErrorAction Stop | Out-Null
                    [system.io.directory]::CreateDirectory($destRemPath) | Out-Null
                    $Acl = Get-Acl $destRemPath  -ErrorAction Stop
                    $Ar = New-Object  system.security.accesscontrol.filesystemaccessrule("$localComputerName$","FullControl", "ContainerInherit,ObjectInherit", "None", "Allow")
                    $Acl.SetAccessRule($Ar) | Out-Null
                    Set-Acl $destRemPath $Acl -ErrorAction Stop | Out-Null

                    $roboCommand = "Robocopy \`"$Source\`" \`"$destRemPath\`" /mir /r:0 /w:0"
                    $scheduledTask = "schtasks /create /ru System /tn '$TaskName' /rl HIGHEST /tr '$roboCommand' /sc $Schedule /MO $Modifier /D $Days /st '$StartTime' /f"
                } else {
                    $localShareName = $officeVersion + "Updates`$"

                    $localShare = Get-SmbShare -Name $localShareName -ErrorAction SilentlyContinue
                    if (!($localShare)) {
                       New-SmbShare -Name "$localShareName" -Path $Source -ReadAccess "Everyone" | Out-Null
                    }
                    $localSharePath = "\\" + $localComputerName + "\" + $localShareName

                    $roboCommand = "Robocopy \`"$localSharePath\`" \`"$destRemPath\`" /mir /r:0 /w:0"
                    $scheduledTask = "schtasks /create /s $serverName /ru System /rl HIGHEST /tn '$TaskName' /tr '$roboCommand' /sc $Schedule /MO $Modifier /D $Days /st '$StartTime' /f"

                    if ($Credential) {
                      $scheduledTask += " /U "
                      $scheduledTask += $Credential.UserName
                      $scheduledTask += " /P "
                      $scheduledTask += $creds.GetNetworkCredential().password
                    }
                }

                $scheduledTaskDel = "schtasks /delete /tn '$TaskName' /f"
        
                try {   
                  & $scheduledTaskDel| out-Null
                } catch { }

                Invoke-Expression $scheduledTask | Out-Null

                $remShare = Get-ODTRemoteUpdateSource | Where { $_.ShareName.ToLower() -eq $remotePath.ToLower() }
                $remShares += $remShare
            } catch {
              Throw
            }
        }  
    
        $remShares 
    }                                          
}

function Disable-ODTRemoteUpdateSourceReplication() {
    Param(
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string[]] $ShareName    
    )

    process {
        $progDirPath = "$env:ProgramFiles\Office Update Replication"
        [system.io.directory]::CreateDirectory($progDirPath) | Out-Null
        $Source = "$progDirPath\Office"

        foreach($remotePath in $ShareName) {
            $serverName= $remotePath.Split("\")[2]
            $shareName= $remotePath.Replace("\\$serverName\", "").Replace("\", "-")

            $TaskName = "Microsoft\OfficeC2R\ODT Replication - $serverName - $shareName"

            $scheduledTaskDel = "schtasks /delete /tn '$TaskName' /f"
           
            Invoke-Expression $scheduledTaskDel | out-Null

            $remShares = Get-ODTRemoteUpdateSource | Where { $_.ShareName.ToLower() -eq $remotePath.ToLower() }
            $remShares
        }       
    }                                      
}


function Add-ODTRemoteUpdateSource() {
   <#

.SYNOPSIS
Create or add a remote share to a list of shares to replicate the ODT
builds to.

.DESCRIPTION
By specifying a file name and list of shares a csv file will be created
containing the list of available shares that will replicate with the source
hosting the C2R builds. If there is not an existing csv file one will be
created.

.PARAMETER ODTShareNameLogFile
The name of the csv file containing a list of shared folders.

.PARAMETER RemoteShares
A list of remote shares to remove from the csv.

.EXAMPLE
Add-ODTRemoteUpdateSource -RemoteShare "\\Computer3\ODT Replication","\\Computer4\ODT Replication" -ODTShareNameLogFile "\\Server1\ODT Replication\ODTRemoteShares.csv"
The ODTRemoteShares.csv file will be updated to include 
shares "\\Computer3\ODT Replication" and "\\Computer4\ODT Replication".

#> 
    Param(
        [Parameter()]
        [string[]] $RemoteShare,

        [Parameter()]
        [OfficeCTRVersion]$OfficeVersion = "Office2013"
    )

    Process {
        $progDirPath = "$env:ProgramFiles\Office Update Replication"
        [system.io.directory]::CreateDirectory($progDirPath) | Out-Null
        $ODTShareNameLogFile = "$progDirPath\ODTReplication.csv"

        [PSObject[]]$ExistingShares = new-object PSObject[] 0;
        if(Test-Path $ODTShareNameLogFile){
            $ExistingShares = Import-Csv $ODTShareNameLogFile
        }

        foreach($share in $RemoteShare) {
            $checkShares = $ExistingShares | Where { $_.ShareName.ToLower() -eq $share.ToLower() }
            if ($checkShares.Length -eq 0) {
                $results = new-object PSObject[] 0;
                $Result = New-Object –TypeName PSObject 

                $computerName = $share.Split("\")[2]

                if (Test-Connection -ComputerName $computerName -ErrorAction SilentlyContinue){
                    if (Test-Path -Path $share) {
                        Add-Member -InputObject $Result -MemberType NoteProperty -Name "ShareName" -Value $Share
                        Add-Member -InputObject $Result -MemberType NoteProperty -Name "AutoReplicationEnabled" -Value $false
                        Add-Member -InputObject $Result -MemberType NoteProperty -Name "OfficeVersion" -Value $OfficeVersion

                        $ExistingShares += $Result

                        $Result
                    } else {
                       Write-Host "The remote share `"$share`" is unavailble" -BackgroundColor Red
                    }
                } else {
                   Write-Host "The remote host `"$computerName`" is unavailble" -BackgroundColor Red
                }
            }
        } 
            
        $ExistingShares | Export-Csv $ODTShareNameLogFile -NoTypeInformation -Force
    }
}

function Remove-ODTRemoteUpdateSource() {
<#

.SYNOPSIS
Remove a remote share from the list of available shares.

.DESCRIPTION
A remote share can be removed from the list of available shares
recorded in the csv file.

.PARAMETER ODTShareNameLogFile
The name of the csv file containing a list of shared folders.

.PARAMETER RemoteShares
A list of remote shares to remove from the csv.

.EXAMPLE
Remove-ODTRemoteUpdateSource -ODTShareNameLogFile "\\Server1\ODT Replication\ODTRemoteShares.csv" -RemoteShares "\\Computer1\ODT Replication","\\Computer2\ODT Replication"
Remote shares \\Computer1\ODT Replication" and "\\Computer2\ODT Replication will
be removed from the csv file and saved.

#>
    [CmdletBinding(SupportsShouldProcess=$True)]
    Param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$true)]
        [string[]] $ShareName
    )

    Process {
        $progDirPath = "$env:ProgramFiles\Office Update Replication"
        [system.io.directory]::CreateDirectory($progDirPath) | Out-Null
        $ODTShareNameLogFile = "$progDirPath\ODTReplication.csv"

        foreach ($remoteShare in $ShareName) {
           Disable-ODTRemoteUpdateSourceReplication -ShareName $remoteShare | Out-Null
        }

        Write-Host "Removing Remote Share: $ShareName"

        $removedShares = Import-Csv $ODTShareNameLogFile | where ShareName -notin $ShareName
        $removedShares | Export-Csv $ODTShareNameLogFile -Force -NoTypeInformation
    }
}

function Get-ODTRemoteUpdateSource() {
<#
.SYNOPSIS
List available shares.

.DESCRIPTION
Given the csv file name the list of available remote shares and their
last updated time will output to the console.

.PARAMETER ODTRemoteUpdateSource
The name of the csv file containing a list of shared folders.

.EXAMPLE
List-ODTRemoteUpdateSource -ODTShareNameLogFile
The csv recording the list of available remote shares will 
be populated in the console.


#>
    [cmdletbinding()]
    Param(
        [Parameter()]
        [OfficeVersionSelection]$OfficeVersion = "All"
    )

    Process {
        $defaultDisplaySet = 'ShareName', 'OfficeVersion', 'AutoReplicationEnabled', 'ReplicationDirection', 'LastReplTime', 'NextReplTime', 'LastResult'

        $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet(‘DefaultDisplayPropertySet’,[string[]]$defaultDisplaySet)
        $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)

        $progDirPath = "$env:ProgramFiles\Office Update Replication"
        [system.io.directory]::CreateDirectory($progDirPath) | Out-Null
        $ODTShareNameLogFile = "$progDirPath\ODTReplication.csv"

        if (!(Test-Path $ODTShareNameLogFile)) {
           return
        }

        $remoteShares = Import-Csv $ODTShareNameLogFile
        if ($OfficeVersion -ne "All") {
           $remoteShares = $remoteShares | Where { $_.OfficeVersion -eq $OfficeVersion }
        }

        $results = new-object PSObject[] 0;

        foreach ($remoteShare in $remoteShares) {
	       $Result = New-Object –TypeName PSObject 
	
           $serverName = $remoteShare.ShareName.Split("\")[2]
           $shareName= $remoteShare.ShareName.Replace("\\$serverName\", "").Replace("\", "-")
       
           $replExists = findReplScheduledTask -ServerName $serverName -ShareName $shareName
           $remoteReplExists = findRemoteReplScheduledTask -ServerName $serverName -ShareName $shareName

           $replDirection = "Push"
           if ($remoteReplExists) {
               $remoteShare.AutoReplicationEnabled = $remoteReplExists
           } else {
                $remoteShare.AutoReplicationEnabled = $replExists
           }

           if ($remoteReplExists) {
              $replDirection = "Pull"
           }
	
           Add-Member -InputObject $Result -MemberType NoteProperty -Name "ShareName" -Value $remoteShare.ShareName
           Add-Member -InputObject $Result -MemberType NoteProperty -Name "AutoReplicationEnabled" -Value $remoteShare.AutoReplicationEnabled
           Add-Member -InputObject $Result -MemberType NoteProperty -Name "ReplicationDirection" -Value $replDirection
           Add-Member -InputObject $Result -MemberType NoteProperty -Name "OfficeVersion" -Value $remoteShare.OfficeVersion
	   
	       $NextRunTime = $null
           $LastRunTime = $null
           $LastResult = 0
           $State = $null
           $ScheduleType= $null
           $StartTime = $null
           $ScheduleDays= $null
           $ScheduleMonths= $null

           if ($replExists -or $replDirection) {
              $TaskName = "Microsoft\OfficeC2R\ODT Replication - $serverName - $shareName"

              if ($replDirection -eq "Push") {
                 schtasks /query /tn $TaskName /v /fo csv  | Out-File -FilePath "$env:temp\TmpSchTask.csv"
              } else {
                 schtasks /query /s "$serverName" /tn $TaskName /v /fo csv  | Out-File -FilePath "$env:temp\TmpSchTask.csv"
              }

	          $importTasks = Import-Csv -Path "$env:temp\TmpSchTask.csv"
	     	
	          $NextRunTime = $importTasks.'Next Run Time'
              $LastRunTime = $importTasks.'Last Run Time'
              $LastResult = $importTasks.'Last Result'
              $ScheduleType = $importTasks.'Schedule Type' 
              $StartTime = $importTasks.'Start Time'
              $ScheduleDays = $importTasks.'Days'
              $ScheduleMonths = $importTasks.'Months'
              $State = $importTasks.'Scheduled Task State'
	       }
	   
	       Add-Member -InputObject $Result -MemberType NoteProperty -Name "NextReplTime" -Value $NextRunTime
           Add-Member -InputObject $Result -MemberType NoteProperty -Name "LastReplTime" -Value $LastRunTime
           Add-Member -InputObject $Result -MemberType NoteProperty -Name "LastResult" -Value $LastResult
           Add-Member -InputObject $Result -MemberType NoteProperty -Name "State" -Value $State
           Add-Member -InputObject $Result -MemberType NoteProperty -Name "StartTime" -Value  $StartTime
           Add-Member -InputObject $Result -MemberType NoteProperty -Name "ScheduleType" -Value $ScheduleType
           Add-Member -InputObject $Result -MemberType NoteProperty -Name "ScheduleDays" -Value  $ScheduleDays
           Add-Member -InputObject $Result -MemberType NoteProperty -Name "ScheduleMonths" -Value $ScheduleMonths
	   
           $Result | Add-Member MemberSet PSStandardMembers $PSStandardMembers

	       $results += $Result
        }

        return $results
    }
}

function Start-OfficeUpdateSourceCleanup() {
    [cmdletbinding()]
    Param(
        [Parameter()]
        [OfficeVersionSelection]$OfficeVersion = "All",

        [Parameter()]
        [int]$NumberOfVersionsToKeep = 2
    )

    Write-Host "Removing Previous Versions"
    Write-Host

    $progDirPath = "$env:ProgramFiles\Office Update Replication"
    [system.io.directory]::CreateDirectory($progDirPath) | Out-Null

    $sourcePaths = @()

    switch($OfficeVersion){
        All { 
           $sourcePaths += "$progDirPath\Office2013\Office\Data" 
           $sourcePaths += "$progDirPath\Office2016\Office\Data"
        }
        Office2013 { $sourcePaths += "$progDirPath\Office2013\Office\Data" }
        Office2016 { $sourcePaths += "$progDirPath\Office2016\Office\Data" }
    }

    foreach ($SourcePath in $sourcePaths) {
        $tempPath = $env:TEMP
        $latestVersionString = $null

        if ($NumberOfVersionsToKeep -lt 1) {
            $NumberOfVersionsToKeep = 1
        }

        if (Test-Path -Path $SourcePath) {
           $v32Path = $SourcePath + "\v32.cab"
           if (Test-Path -Path $v32Path) {
              expand $v32Path $tempPath -f:VersionDescriptor.xml | Out-Null
              $xmlPath = $tempPath + "\VersionDescriptor.xml"
              [xml]$xmlVersion = Get-Content $xmlPath
              $buildVersion = $xmlVersion.Version.Available.Build
              $latestVersionString = $buildVersion
              Remove-Item -Path $xmlPath
           }
        }

        if ($latestVersionString) {
           $latestVerion = New-Object -TypeName System.Version -ArgumentList @($latestVersionString)
      
           $versionFolders = Get-ChildItem -Path $SourcePath -Include $include -Recurse:$recurse | Where-Object { $_.PSIsContainer } | select Name
           $sortedFolders = $versionFolders | Sort-Object -Descending -Property Name

           Write-Host "Checking Path: $SourcePath"

           for ($n=$NumberOfVersionsToKeep;$n -lt $sortedFolders.Length;$n++) {
             try {
              $folder = $sortedFolders[$n]
              $folderName = $folder.Name

              Write-Host "`tRemoving Version: $folderName"

              if (Test-Path -Path "$SourcePath\v32_$folderName.cab") {
                 Remove-Item -Path "$SourcePath\v32_$folderName.cab" -Force -ErrorAction Stop | Out-Null
              }

              if (Test-Path -Path "$SourcePath\v64_$folderName.cab") {
                 Remove-Item -Path "$SourcePath\v64_$folderName.cab" -Force -ErrorAction Stop | Out-Null
              }

              if (Test-Path -Path "$SourcePath\$folderName") {
                 Remove-Item -Path "$SourcePath\$folderName" -Recurse -Force -ErrorAction Stop
              }
            } catch {
              Throw
            }
           }

           Write-Host
        }
    }
}


function Copy-WithProgress {  
    [CmdletBinding()]  
  
    param (  
            [Parameter(Mandatory = $true)]  
            [string] $Source  
        , [Parameter(Mandatory = $true)]  
            [string] $Destination)  
  
    $robocopycmd = "robocopy ""$source"" ""$destination"" /mir /bytes"  
    $Staging = Invoke-Expression "$robocopycmd /l"  
    $totalnewfiles = $Staging -match 'new file'  
    $totalmodified = $Staging -match 'newer'  
    $totalfiles = $totalnewfiles + $totalmodified 
    $TotalBytesarray = @() 
    foreach ($file in $totalfiles)   
    {  
        $fileSize = getFileSize -text $file
        $TotalBytesarray+=$fileSize
    }  
    $totalbytes = (($TotalBytesarray | Measure-Object -Sum).sum) 
  
    $robocopyjob = Start-Job -Name robocopy -ScriptBlock {param ($command) ; Invoke-Expression -Command $command} -ArgumentList $robocopycmd  
  
    while ($robocopyjob.State -eq 'running')  
    {  
        $progress = Receive-Job -Job $robocopyjob -Keep -ErrorAction SilentlyContinue 
        if ($progress) 
        { 
            $copiedfiles = ($progress | Select-String -SimpleMatch 'new file', 'newer') 
            if ($copiedfiles.count -le 0) { $TotalFilesCopied = $copiedfiles.Count } 
            else { $TotalFilesCopied = $copiedfiles.Count - 1 } 
            $FilesRemaining = ($totalfiles.count - $TotalFilesCopied) 
            $Bytesarray = @() 
            foreach ($Newfile in $copiedfiles) 
            { 
                $fileSize = getFileSize -text $Newfile
                $Bytesarray+=$fileSize
            } 
            $bytescopied = ([int64]$Bytesarray[-1] * ($Filepercentcomplete/100)) 
            $totalfilebytes = [int64]$Bytesarray[-1] 
            $TotalBytesCopied = ((($Bytesarray | Measure-Object -Sum).sum) - $totalfilebytes) + $bytescopied 
            $TotalBytesRemaining = ($totalbytes - $totalBytesCopied) 
            if ($copiedfiles) 
            { 
                #$fileSize = getFileSize -text $copiedfiles[-1].tostring()

                $currentfile = getFileName -text $copiedfiles[-1].tostring()
            } 
            $totalfilescount = $totalfiles.count 
            if ($progress[-1] -match '%') { $Filepercentcomplete = $progress[-1].substring(0, 3).trim() } 
            else { $Filepercentcomplete = 0 } 
            $totalPercentcomplete = (($TotalBytesCopied/$totalbytes) * 100) 
            if ($totalbytes -gt 2gb) { $BytesCopiedprogress = "{0:N2}" -f ($totalBytesCopied/1gb); $totalbytesprogress = "{0:N2}" -f ($totalbytes/1gb); $bytes = 'Gbytes' } 
            else { $BytesCopiedprogress = "{0:N2}" -f ($totalBytesCopied/1mb); $totalbytesprogress = "{0:N2}" -f ($totalbytes/1mb); $bytes = 'Mbytes' } 
            if ($totalfilebytes -gt 1gb) { $totalfilebytes = "{0:N2}" -f ($totalfilebytes/1gb); $bytescopied = "{0:N2}" -f ($bytescopied/1gb); $filebytes = 'Gbytes' } 
            else { $totalfilebytes = "{0:N2}" -f ($totalfilebytes/1mb); $bytescopied = "{0:N2}" -f ($bytescopied/1mb); $filebytes = 'Mbytes' } 
             
            if ($currentfile) {
               Write-Progress -Id 1 -Activity "Copying files from $source to $destination, $totalfilescopied of $totalfilescount files copied" -Status "$bytescopiedprogress of $totalbytesprogress $bytes copied" -PercentComplete $totalPercentcomplete 
               Write-Progress -Id 2 -Activity "$currentfile" -status "$bytescopied of $totalfilebytes $filebytes" -PercentComplete $Filepercentcomplete 
            }
        } 
         
    } 
     
    Write-Progress -Id 1 -Activity "Copying files from $source to $destination" -Status 'Completed' -Completed  
    Write-Progress -Id 2 -Activity 'Done' -Completed  
    $results = Receive-Job -Job $robocopyjob  
    Remove-Job $robocopyjob  
    $results[5]  
    $results[-13..-1]  
} 

function findScheduledTask() {
    param(
        [OfficeCTRVersion]$OfficeVersion = "Office2013",
        [string]$Bitness="32"
    )  
    
     $TaskName = "Microsoft\OfficeC2R\$OfficeVersion ODT Download $Bitness-bit"
     $scheduledTaskQuery = "/query /tn `"$TaskName`""
 
        $pinfo = New-Object System.Diagnostics.ProcessStartInfo
        $pinfo.FileName = "schtasks"
        $pinfo.RedirectStandardError = $true
        $pinfo.RedirectStandardOutput = $true
        $pinfo.UseShellExecute = $false
        $pinfo.Arguments = $scheduledTaskQuery

        $p = New-Object System.Diagnostics.Process
        $p.StartInfo = $pinfo
        $p.Start() | Out-Null
        $p.WaitForExit()
        $stdout = $p.StandardOutput.ReadToEnd()
        $stderr = $p.StandardError.ReadToEnd()

        if ($stderr) {
            return $false;
        }

        if ($stdout) {
            return $true
        }

     return $false
}

function findReplScheduledTask() {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ServerName,
        [Parameter(Mandatory=$true)]
        [string]$ShareName
    )  
    
     $TaskName = "Microsoft\OfficeC2R\ODT Replication - $ServerName - $ShareName"
     $scheduledTaskQuery = "/query /tn `"$TaskName`""
 
        $pinfo = New-Object System.Diagnostics.ProcessStartInfo
        $pinfo.FileName = "schtasks"
        $pinfo.RedirectStandardError = $true
        $pinfo.RedirectStandardOutput = $true
        $pinfo.UseShellExecute = $false
        $pinfo.Arguments = $scheduledTaskQuery
        $pinfo.CreateNoWindow = $true

        $p = New-Object System.Diagnostics.Process
        $p.StartInfo = $pinfo
        $p.Start() | Out-Null
        $p.WaitForExit()
        $stdout = $p.StandardOutput.ReadToEnd()
        $stderr = $p.StandardError.ReadToEnd()

        if ($stderr) {
            return $false;
        }

        if ($stdout) {
            return $true
        }

     return $false
}

function findRemoteReplScheduledTask() {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ServerName,
        [Parameter(Mandatory=$true)]
        [string]$ShareName
    )  
    
     $TaskName = "Microsoft\OfficeC2R\ODT Replication - $ServerName - $ShareName"
     $scheduledTaskQuery = "/query /s $ServerName /tn `"$TaskName`""
 
        $pinfo = New-Object System.Diagnostics.ProcessStartInfo
        $pinfo.FileName = "schtasks"
        $pinfo.RedirectStandardError = $true
        $pinfo.RedirectStandardOutput = $true
        $pinfo.UseShellExecute = $false
        $pinfo.Arguments = $scheduledTaskQuery
        $pinfo.CreateNoWindow = $true

        $p = New-Object System.Diagnostics.Process
        $p.StartInfo = $pinfo
        $p.Start() | Out-Null
        $p.WaitForExit()
        $stdout = $p.StandardOutput.ReadToEnd()
        $stderr = $p.StandardError.ReadToEnd()

        if ($stderr) {
            return $false;
        }

        if ($stdout) {
            return $true
        }

     return $false
}

function getSecondTuesday() {
    $FindNthDay=2

    $WeekDay='Tuesday'

    [datetime]$Today=[datetime]::NOW
    $todayM=$Today.AddMonths(1).Month.ToString()
    $todayY=$Today.Year.ToString()
    [datetime]$StrtMonth=$todayM+'/1/'+$todayY

    while ($StrtMonth.DayofWeek -ine $WeekDay ) { $StrtMonth=$StrtMonth.AddDays(1) }

    $secTues = $StrtMonth.AddDays(7*($FindNthDay-1))
    
    $dateReturn = Get-Date -Year $secTues.Year -Month $secTues.Month -Day $secTues.Day -Hour 18 -Minute 00 
    return $dateReturn 
}

function getFileName() {
    param (  
            [Parameter(Mandatory = $true)]  
            [string] $text 
    )

    $text = $text -replace "`t", " "

    foreach ($line in $text.Split(" ")) {
       if ($line) {
          $line = $line.Trim()
          if ($line.Length -gt 0) {
              if ($line.Contains(".")) {
                 return $line
              }
          }
       }
    }

}

function getFileSize() {
    param (  
            [Parameter(Mandatory = $true)]  
            [string] $text 
    )

    $text = $text -replace "`t", " "

    foreach ($line in $text.Split(" ")) {
       if ($line) {
          $line = $line.Trim()
          if ($line.Length -gt 0) {
              if (IsNumeric -Value $line) {
                 return $line
              }
          }
       }
    }

}

function IsNumeric { 
 
[CmdletBinding( 
    SupportsShouldProcess=$True, 
    ConfirmImpact='High')] 
 
param ( 
 
[Parameter( 
    Mandatory=$True, 
    ValueFromPipeline=$True, 
    ValueFromPipelineByPropertyName=$True)] 
     
    $Value, 
     
[Parameter( 
    Mandatory=$False, 
    ValueFromPipeline=$True, 
    ValueFromPipelineByPropertyName=$True)] 
    [alias('B')] 
    [Switch] $Boolean 
     
) 
     
BEGIN { 
 
    #clear variable 
    $IsNumeric = 0 
 
} 
 
PROCESS { 
 
    #verify input value is numeric data type 
    try { 0 + $Value | Out-Null 
    $IsNumeric = 1 }catch{ $IsNumeric = 0 } 
 
    if($IsNumeric){  
        $IsNumeric = 1 
        if($Boolean) { $Isnumeric = $True } 
    }else{  
        $IsNumeric = 0 
        if($Boolean) { $IsNumeric = $False } 
    } 
     
    if($PSBoundParameters['Verbose'] -and $IsNumeric) {  
    Write-Verbose "True" }else{ Write-Verbose "False" } 
     
    
    return $IsNumeric 
} 
 
END {} 
 
} 

Function Set-ODTAdd{
<#
.SYNOPSIS
Modifies an existing configuration xml file's add section

.PARAMETER SourcePath
Optional.
The SourcePath value can be set to a network, local, or HTTP path that contains a 
Click-to-Run source. Environment variables can be used for network or local paths.
SourcePath indicates the location to save the Click-to-Run installation source 
when you run the Office Deployment Tool in download mode.
SourcePath indicates the installation source path from which to install Office 
when you run the Office Deployment Tool in configure mode. If you don’t specify 
SourcePath in configure mode, Setup will look in the current folder for the Office 
source files. If the Office source files aren’t found in the current folder, Setup 
will look on Office 365 for them.
SourcePath specifies the path of the Click-to-Run Office source from which the 
App-V package will be made when you run the Office Deployment Tool in packager mode.
If you do not specify SourcePath, Setup will attempt to create an \Office\Data\... 
folder structure in the working directory from which you are running setup.exe.

.PARAMETER Version
Optional. If a Version value is not set, the Click-to-Run product installation streams 
the latest available version from the source. The default is to use the most recently 
advertised build (as defined in v32.CAB or v64.CAB at the Click-to-Run Office installation source).
Version can be set to an Office 2013 build number by using this format: X.X.X.X

.PARAMETER Bitness
Required. Specifies the edition of Click-to-Run for Office 365 product to use: 32- or 64-bit.

.PARAMETER TargetFilePath
Full file path for the file to be modified and be output to.

.Example
Set-ODTAdd -SourcePath "C:\Preload\Office" -TargetFilePath "$env:Public/Documents/config.xml"
Sets config SourcePath property of the add element to C:\Preload\Office

.Example
Set-ODTAdd -SourcePath "C:\Preload\Office" -Version "15.1.2.3" -TargetFilePath "$env:Public/Documents/config.xml"
Sets config SourcePath property of the add element to C:\Preload\Office and version to 15.1.2.3

.Notes
Here is what the portion of configuration file looks like when modified by this function:

<Configuration>
  ...
  <Add SourcePath="\\server\share\" Version="15.1.2.3" OfficeClientEdition="32"> 
      ...
  </Add>
  ...
</Configuration>

#>
    Param(

        [Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true, Position=0)]
        [string] $ConfigurationXML = $NULL,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $SourcePath = $NULL,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $Version,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $Bitness,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $TargetFilePath

    )

    Process{
        $TargetFilePath = GetFilePath -TargetFilePath $TargetFilePath

        #Load file
        [System.XML.XMLDocument]$ConfigFile = New-Object System.XML.XMLDocument

        if ($TargetFilePath) {
           $ConfigFile.Load($TargetFilePath) | Out-Null
        } else {
            if ($ConfigurationXml) 
            {
              $ConfigFile.LoadXml($ConfigurationXml) | Out-Null
              $global:saveLastConfigFile = $NULL
              $global:saveLastFilePath = $NULL
            }
        }

        $global:saveLastConfigFile = $ConfigFile.OuterXml

        #Check for proper root element
        if($ConfigFile.Configuration -eq $null){
            throw $NoConfigurationElement
        }

        #Get Add element if it exists
        if($ConfigFile.Configuration.Add -eq $null){
            [System.XML.XMLElement]$AddElement=$ConfigFile.CreateElement("Add")
            $ConfigFile.Configuration.appendChild($AddElement) | Out-Null
        }

        #Set values as desired
        if([string]::IsNullOrWhiteSpace($SourcePath) -eq $false){
            $ConfigFile.Configuration.Add.SetAttribute("SourcePath", $SourcePath) | Out-Null
        } else {
            if ($PSBoundParameters.ContainsKey('SourcePath')) {
                $ConfigFile.Configuration.Add.RemoveAttribute("SourcePath")
            }
        }

        if([string]::IsNullOrWhiteSpace($Version) -eq $false){
            $ConfigFile.Configuration.Add.SetAttribute("Version", $Version) | Out-Null
        } else {
            if ($PSBoundParameters.ContainsKey('Version')) {
                $ConfigFile.Configuration.Add.RemoveAttribute("Version")
            }
        }

        if([string]::IsNullOrWhiteSpace($Bitness) -eq $false){
            $ConfigFile.Configuration.Add.SetAttribute("OfficeClientEdition", $Bitness) | Out-Null
        } else {
            if ($PSBoundParameters.ContainsKey('OfficeClientEdition')) {
                $ConfigFile.Configuration.Add.RemoveAttribute("OfficeClientEdition")
            }
        }

        $ConfigFile.Save($TargetFilePath) | Out-Null
        $global:saveLastFilePath = $TargetFilePath

        if (($PSCmdlet.MyInvocation.PipelineLength -eq 1) -or `
            ($PSCmdlet.MyInvocation.PipelineLength -eq $PSCmdlet.MyInvocation.PipelinePosition)) {
            Write-Host

            Format-XML ([xml](cat $TargetFilePath)) -indent 4

            Write-Host
            Write-Host "The Office XML Configuration file has been saved to: $TargetFilePath"
        } else {
            $results = new-object PSObject[] 0;
            $Result = New-Object –TypeName PSObject 
            Add-Member -InputObject $Result -MemberType NoteProperty -Name "TargetFilePath" -Value $TargetFilePath
            Add-Member -InputObject $Result -MemberType NoteProperty -Name "SourcePath" -Value $SourcePath
            Add-Member -InputObject $Result -MemberType NoteProperty -Name "Version" -Value $Version
            Add-Member -InputObject $Result -MemberType NoteProperty -Name "Bitness" -Value $Bitness
            $Result
        }
    }

}

Function GetFilePath() {
    Param(
       [Parameter(ValueFromPipelineByPropertyName=$true)]
       [string] $TargetFilePath
    )

    if (!($TargetFilePath)) {
        $TargetFilePath = $global:saveLastFilePath
    }  

    if (!($TargetFilePath)) {
       Write-Host "Enter the path to the XML Configuration File: " -NoNewline
       $TargetFilePath = Read-Host
    } else {
       #Write-Host "Target XML Configuration File: $TargetFilePath"
    }

    return $TargetFilePath
}