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

function Download-ODTOfficeFiles() {
    param(
        [OfficeCTRVersion]$OfficeVersion,
        [string] $XmlConfigPath,
        [string] $TaskName = $null,
        [string[]] $ComputerName = $env:COMPUTERNAME
    )  
    
    switch($OfficeVersion){

            Office2013 { $XmlDownload = ".\Office2013Setup.exe /Download $XmlConfigPath" }
            Office2016 { $XmlDownload = ".\Office2016Setup.exe /Download $XmlConfigPath" }
        }

    if(!($TaskName)){
    
        Invoke-Expression $XmlDownload
    }
    else{

        foreach($computer in $ComputerName){
            
            $scheduledTask = 'schtasks /create /s $computer /ru System /tn $TaskName /tr $XmlDownload /sc Daily /st 03:00:00'

            Invoke-Expression $scheduledTask
        }
    }
}

function Replicate-ODTOfficeFiles() {

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
    
    Param(
        [string[]] $RemoteShare,
        [string] $ODTShareNameLogFile
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

    Param(
        [string] $ODTShareNameLogFile,
        [string[]] $RemoteShares
    )

    $removedShares = Import-Csv $ODTShareNameLogFile | where ShareName -notin $RemoteShares
    $removedShares | Export-Csv $ODTShareNameLogFile -Force -NoTypeInformation
}

function List-ODTRemoteUpdateSource() {

    Param(
        [string] $ODTShareNameLogFile
    )

    Import-Csv $ODTShareNameLogFile
}