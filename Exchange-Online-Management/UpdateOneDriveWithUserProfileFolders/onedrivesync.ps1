$dirs = Get-ChildItem $env:USERPROFILE
$oneDriveFolder = ""
foreach($dir in $dirs){
    if($dir.ToString().toLower() -match "onedrive" -and ($dir.ToString() -match "-" -or $dir.ToString().toLower() -match "for business")){        
        $oneDriveFolder = $env:USERPROFILE + "\" + $dir.ToString()
        Write-Host $oneDriveFolder
    }
}

$UsersFolders = New-Object "System.Collections.Generic.List[String]"

$UsersFolders += $env:USERPROFILE + "\" + "Desktop"
#$UsersFolders += $env:USERPROFILE + "\" + "Documents"
#$UsersFolders += $env:USERPROFILE + "\" + "Downloads"
#$UsersFolders += $env:USERPROFILE + "\" + "Music"
$UsersFolders += $env:USERPROFILE + "\" + "Pictures"

foreach($userFolder in $UsersFolders){
    $destination = $oneDriveFolder + $userFolder.Substring($userFolder.LastIndexOf("\"))
    robocopy $userFolder $destination
}