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
      Office2013,Office2016
   }
"@ 
Add-Type -TypeDefinition $OfficeCTRVersion

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
        [OfficeCTRVersion]$OfficeVersion = "Office2013",
        [string] $XmlConfigPath = "$PSScriptRoot\configuration.xml"
    )  

    . $PSScriptRoot\Edit-OfficeConfigurationFile.ps1
    
    switch($OfficeVersion){
       Office2013 { $odtExtPath = "$PSScriptRoot\Office2013Setup.exe" }
       Office2016 { $odtExtPath = "$PSScriptRoot\Office2016Setup.exe" }
    }

    $progDirPath = "$env:ProgramFiles\OfficeCTRRepl"
    [system.io.directory]::CreateDirectory($progDirPath) | Out-Null

    Write-Host "Downloading `"$OfficeVersion`" Latest 32-Bit Version..." -NoNewline
    $download32 = "$odtExtPath /download $XmlConfigPath"
    Set-ODTAdd -TargetFilePath "$XmlConfigPath" -Bitness 32 -SourcePath $progDirPath | Out-Null
    Invoke-Expression $download32
    Write-Host "Completed"

    Write-Host "Downloading `"$OfficeVersion`" Latest 64-Bit Version..." -NoNewline
    $download64 = "$odtExtPath /download $XmlConfigPath"
    Set-ODTAdd -TargetFilePath "$XmlConfigPath" -Bitness 64 -SourcePath $progDirPath | Out-Null
    Invoke-Expression $download64
    Write-Host "Completed"
}

function New-ODTDownloadSchedule() {
    param(
        [OfficeCTRVersion]$OfficeVersion = "Office2013",
        [string] $XmlConfigPath = "$PSScriptRoot\configuration.xml",
        [string] $ScheduledTime32Bit = "19:00",
        [string] $ScheduledTime64Bit = "18:00"
    )  
    
    $progDirPath = "$env:ProgramFiles\OfficeCTRRepl"
    [system.io.directory]::CreateDirectory($progDirPath) | Out-Null

    switch($OfficeVersion){
       Office2013 { 
         if (!(Test-Path -Path "$env:ProgramFiles\OfficeCTRRepl\Office2013Setup.exe")) {
            Copy-Item -Path "$PSScriptRoot\Office2013Setup.exe" -Destination "$env:ProgramFiles\OfficeCTRRepl\Office2013Setup.exe" -Force | Out-Null
         }
       }
       Office2016 { 
          if (!(Test-Path -Path "$env:ProgramFiles\OfficeCTRRepl\Office2016Setup.exe")) {
            Copy-Item -Path "$PSScriptRoot\Office2016Setup.exe" -Destination "$env:ProgramFiles\OfficeCTRRepl\Office2016Setup.exe" -Force -ErrorAction SilentlyContinue | Out-Null
          }
       }
    }

    Copy-Item -Path $XmlConfigPath -Destination "$env:ProgramFiles\OfficeCTRRepl\configuration32.xml" -Force | Out-Null
    Copy-Item -Path $XmlConfigPath -Destination "$env:ProgramFiles\OfficeCTRRepl\configuration64.xml" -Force | Out-Null

    Set-ODTAdd -TargetFilePath "$env:ProgramFiles\OfficeCTRRepl\configuration32.xml" -SourcePath $progDirPath -Bitness 32 | Out-Null
    Set-ODTAdd -TargetFilePath "$env:ProgramFiles\OfficeCTRRepl\configuration64.xml" -SourcePath $progDirPath -Bitness 64 | Out-Null

    switch($OfficeVersion){
       Office2013 { 
            $odtCmd32 = "\`"$progDirPath\Office2013Setup.exe\`" /Download \`"$env:ProgramFiles\OfficeCTRRepl\configuration32.xml\`"" 
            $odtCmd64 = "\`"$progDirPath\Office2013Setup.exe\`" /Download \`"$env:ProgramFiles\OfficeCTRRepl\configuration64.xml\`"" 
       }
       Office2016 { 
            $odtCmd32 = "\`"$progDirPath\Office2016Setup.exe\`" /Download \`"$env:ProgramFiles\OfficeCTRRepl\configuration32.xml\`"" 
            $odtCmd64 = "\`"$progDirPath\Office2016Setup.exe\`" /Download \`"$env:ProgramFiles\OfficeCTRRepl\configuration64.xml\`"" 
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

function Remove-ODTDownloadSchedule() {
    param(
        [OfficeCTRVersion]$OfficeVersion = "Office2013"
    )  
    
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


function Replicate-ODTOfficeFiles() {
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
        [string[]] $Source,
        [string[]] $ODTShareNameLogFile
    )

    [array]$ShareName = Import-Csv $ODTShareNameLogFile | foreach {$_.ShareName}

    foreach($share in $ShareName){

        $destinationFolder = Get-ChildItem $share -Recurse
        $sourceFolder = Get-ChildItem $Source -Recurse

        if($destinationFolder -ne $null){          
            $comparison = Compare-Object -ReferenceObject $sourceFolder -DifferenceObject $destinationFolder -IncludeEqual
            $roboCopy = "Robocopy $source $share /e /np"

            if($comparison.SideIndicator -eq "<="){

                Invoke-Expression $roboCopy
            }
            elseif($comparison.SideIndicator -eq "=="){

                Write-Host "The folders are up to date in $share"
            }
        }
        elseif($destinationFolder -eq $null){
             
            $roboCopy = "Robocopy $source $share /e /np"

            Invoke-Expression $roboCopy
        }                         
    }
}


function Schedule-ODTRemoteShareReplicationTask{
<#
.SYNOPSIS
Create a scheduled task on the remote computer to copy the 
C2R folders from the source on a monthly schedule.

.DESCRIPTION
Given a computer name, source, taskname and the necessary commands
for the task to operate (Schedule,Modifier,Days,StartTime) a scheduled
task can be created on the remote computers to copy the files from
the source.

.PARAMETER ComputerName
LIst of computers to create the shceduled task on.

.PARAMETER Source
The source share hosting the C2R builds.

.PARAMETER TaskName
The name of the scheduled task.

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
    Param(
        [string[]] $ComputerName = $env:COMPUTERNAME,
        [string] $Source,
        [string] $TaskName,
        [Schedule] $Schedule,
        [Modifier] $Modifier,
        [Days] $Days,
        [string] $StartTime = $null         
    )

    foreach($Computer in $ComputerName){

        $Destination = Read-Host "Enter the remote share for $computer"
        $roboCommand = "Robocopy $Source $Destination /e /np"
        $scheduledTask = 'schtasks /create /s $Computer /ru System /tn $TaskName /tr $roboCommand /sc $Schedule /MO $Modifier /D $Days /st $StartTime'   
            
        Invoke-Expression $scheduledTask
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
        [string[]] $RemoteShare,
        [string] $ODTShareNameLogFile = "$env:WinDir\Temp\ODTReplication.csv"
    )
        
    if(!(Test-Path $ODTShareNameLogFile)){

        [array]$RemoteShareTable = foreach($share in $RemoteShare){
       
            $LastWriteTime = Get-ItemProperty $share | foreach {$_.LastWriteTime.ToString("MM-dd-yyyy")}
            $results = new-object PSObject[] 0;
            $Result = New-Object –TypeName PSObject 

            Add-Member -InputObject $Result -MemberType NoteProperty -Name "ShareName" -Value $Share
            Add-Member -InputObject $Result -MemberType NoteProperty -Name "LastUpdateTime" -Value $LastWriteTime

            $result
        } 

        $RemoteShareTable | Export-Csv $ODTShareNameLogFile -NoTypeInformation
    }
    else{

        [array] $AddNewShare = foreach($share in $RemoteShare){

            $LastWriteTime = Get-ItemProperty $share | foreach {$_.LastWriteTime.ToString("MM-dd-yyyy")}
            $results = new-object PSObject[] 0;
            $Result = New-Object –TypeName PSObject 
            Add-Member -InputObject $Result -MemberType NoteProperty -Name "ShareName" -Value $Share
            Add-Member -InputObject $Result -MemberType NoteProperty -Name "LastUpdateTime" -Value $LastWriteTime

            $result
        }

        [array]$ExistingShares = Import-Csv $ODTShareNameLogFile
        $newShares = $ExistingShares += [array]$AddNewShare 
        $newShares | Export-Csv $ODTShareNameLogFile -NoTypeInformation -Force
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
    Param(
        [string] $ODTShareNameLogFile = "$env:WinDir\Temp\ODTReplication.csv",
        [string[]] $RemoteShare
    )

    $removedShares = Import-Csv $ODTShareNameLogFile | where ShareName -notin $RemoteShare
    $removedShares | Export-Csv $ODTShareNameLogFile -Force -NoTypeInformation
}

function List-ODTRemoteUpdateSource() {
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
    Param(
       [string] $ODTShareNameLogFile = "$env:WinDir\Temp\ODTReplication.csv"
    )

    Import-Csv $ODTShareNameLogFile
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