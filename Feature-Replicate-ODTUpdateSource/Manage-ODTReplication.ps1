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

function Add-ODTRemoteUpdateSource() {
    param(
        [string[]]$RemoteShares,
        [string[]]$Languages,
        [string] $ODTShareNameLogFile
    )
   
    $defaultDisplaySet = 'ShareName';
    $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet(‘DefaultDisplayPropertySet’,[string[]]$defaultDisplaySet);
    $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet);

    $results = New-Object PSObject[] 1;

    if(!(Test-Path $ODTShareNameLogFile)){
        foreach($share in $RemoteShares){
            $object = New-Object PSObject -Property @{'ShareName' = "$share";}
            $object | Add-Member MemberSet PSStandardMembers $PSStandardMembers;
            $results += $object
            $results | Export-Csv $ODTShareNameLogFile -ErrorAction SilentlyContinue -NoTypeInformation
        }
    }
    else{
        foreach($share in $RemoteShares){
            $logContent = Import-Csv $ODTShareNameLogFile -Header "ShareName"
            $newRow = New-Object PSObject -Property @{ShareName = $share}
            $logContent += $newRow | Export-Csv $ODTShareNameLogFile -Append -Force -NoTypeInformation
        }
    }
}

function Remove-ODTRemoteUpdateSource() {

    Param(
        [string] $ODTShareNameLogFile,
        [string] $RemoteShares
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