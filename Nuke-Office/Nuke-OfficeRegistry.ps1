function Get-TaskSubFolders {                        
    [cmdletbinding()]                        
    param (                        
        $FolderRef                        
    )                        
    $ArrFolders = @()                        
    $folders = $folderRef.getfolders(1)                        
    if($folders) {                        
        foreach ($folder in $folders) {                        
            $ArrFolders = $ArrFolders + $folder                        
            if($folder.getfolders(1)) {                        
                Get-TaskSubFolders -FolderRef $folder                        
            }                        
        }                        
    }                        
    return $ArrFolders                        
} 

[string] $TaskName = "OfficeRegClean"

$mytask
                     
$SchService = New-Object -ComObject Schedule.Service                        
$SchService.Connect($Computer)                        
$Rootfolder = $SchService.GetFolder("\")            
$folders = @($RootFolder)             
$folders += Get-Tasksubfolders -FolderRef $RootFolder
                                
foreach($Folder in $folders) {                        
    $Tasks = $folder.gettasks(1)                        
    foreach($Task in $Tasks) {           
        if($Task.TaskName -eq $TaskName){
            $preErrorAction = $ErrorActionPreference
            $deleteTask = "schtasks.exe /delete /s $env:COMPUTERNAME /tn $TaskName /F"
            $ErrorActionPreference = Stop
            Invoke-Expression $deleteTask
            $ErrorActionPreference = $preErrorAction

            $Hives = Get-ChildItem Microsoft.PowerShell.Core\Registry::

            $OfficeRegistries = foreach($Hive in $Hives){
                Get-ChildItem "$($Hive.PSPath)" -Recurse -ErrorAction SilentlyContinue | ? PSPath -like *\Software\Microsoft\Office
            }
            foreach($Item in $OfficeRegistries){
                Remove-Item $Item.PSPath -Recurse -Force -ErrorAction SilentlyContinue
            }
        }                                               
    }                        
}