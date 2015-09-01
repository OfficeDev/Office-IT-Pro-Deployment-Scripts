Add-Type -TypeDefinition @"
   public enum OfficeCTRVersion
   {
      Office2013
   }
"@ 

function Download-ODTOfficeFiles() {
    param(
        [OfficeCTRVersion]$OfficeVersion,
        [string] $XmlConfigPath
    )  
    
    switch($OfficeVersion){

        Office2013 { $XmlDownload = "Office2013Setup.exe /Download $XmlConfigPath" }
        Office2016 { $XmlDownload = "Office2016Setup.exe /Download $XmlConfigPath" }
    }

    Invoke-Expression $XmlDownload
}

function Replicate-ODTOfficeFiles() {
   param(
     [bool]$CopyLatestVersionOnly = $true,
     [string] $LogPath,
     [string] $source
   )   

    $Shares = @(Import-Csv $logPath | foreach {$_.ShareName})

    foreach($share in $Shares){

        $roboCopy = "Robocopy $source $share /e /np"

        Invoke-Expression $roboCopy
    }
}

function Add-ODTRemoteUpdateSource() {
    param(
        [string[]]$RemoteShares,
        [string[]]$Languages,
        [string] $LogPath
    )
   
    $defaultDisplaySet = 'ShareName';
    $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet(‘DefaultDisplayPropertySet’,[string[]]$defaultDisplaySet);
    $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet);

    $results = New-Object PSObject[] 1;

    if(!(Test-Path $LogPath)){
        foreach($share in $RemoteShares){
            $object = New-Object PSObject -Property @{'ShareName' = "$share";}
            $object | Add-Member MemberSet PSStandardMembers $PSStandardMembers;
            $results += $object
            $results | Export-Csv $LogPath -ErrorAction SilentlyContinue -NoTypeInformation
        }
    }
    else{
        foreach($share in $RemoteShares){
            $logContent = Import-Csv $LogPath -Header "ShareName"
            $newRow = New-Object PSObject -Property @{ShareName = $share}
            $logContent += $newRow | Export-Csv $LogPath -ErrorAction SilentlyContinue -NoTypeInformation
        }
    }
}

function Remove-ODTRemoteUpdateSource() {

    Param(
        [string] $logPath,
        [string] $newLogPath,
        [string] $shareName
    )

    $files = Import-Csv $logPath | where {$_.ShareName -notlike $shareName}

    $output = @($files)
    $output = $files
    $output
}

function List-ODTRemoteUpdateSource() {



}