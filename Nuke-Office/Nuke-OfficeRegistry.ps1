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
        $mytask = New-Object -TypeName PSobject                         
        $mytask | Add-Member -MemberType NoteProperty -Name ComputerName -Value $Computer                        
        $mytask | Add-Member -MemberType NoteProperty -Name TaskName -Value $Task.Name                        
        $mytask | Add-Member -MemberType NoteProperty -Name TaskFolder -Value $Folder.path                        
        $mytask | Add-Member -MemberType NoteProperty -Name IsEnabled -Value $task.enabled                        
        $mytask | Add-Member -MemberType NoteProperty -Name LastRunTime -Value $task.LastRunTime                        
        $mytask | Add-Member -MemberType NoteProperty -Name NextRunTime -Value $task.NextRunTime
        if($mytask.Name -eq $TaskName){
            $preErrorAction = $ErrorActionPreference
            $deleteTask = "schtasks.exe /delete /s $computer /tn $TaskName /F"
            $ErrorActionPreference = Stop
            Invoke-Expression $deleteTask
            $ErrorActionPreference = $preErrorAction
            break;
        }                                               
    }                        
}

$Hives = Get-ChildItem Microsoft.PowerShell.Core\Registry::

$OfficeRegistries = foreach($Hive in $Hives){
    Get-ChildItem "$($Hive.PSPath)" -Recurse -ErrorAction SilentlyContinue | ? PSPath -like *\Software\Microsoft\Office
}
foreach($Item in $OfficeRegistries){
    Remove-Item $Item.PSPath -Recurse -Force -ErrorAction SilentlyContinue
}