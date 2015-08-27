function Nuke-Office{

[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [string[]]$ComputerName,
    [switch]$ShowAllInstalledProducts,
    [System.Management.Automation.PSCredential]$Credentials
)

function Get-OfficeVersion{

[CmdletBinding(SupportsShouldProcess=$true)]
    param(
        [Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true, Position=0)]
        [string[]]$ComputerName,
        [switch]$ShowAllInstalledProducts,
        [System.Management.Automation.PSCredential]$Credentials
    )

    begin {
        $HKLM = [UInt32] "0x80000002"
        $HKCR = [UInt32] "0x80000000"

        $excelKeyPath = "Excel\DefaultIcon"
        $wordKeyPath = "Word\DefaultIcon"
   
        $installKeys = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
                       'SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall'

        $officeKeys = 'SOFTWARE\Microsoft\Office',
                      'SOFTWARE\Wow6432Node\Microsoft\Office'

        $defaultDisplaySet = 'ComputerName','DisplayName','Version', 'ClicktoRun'

        $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet(‘DefaultDisplayPropertySet’,[string[]]$defaultDisplaySet)
        $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
    }


    process {

     $results = new-object PSObject[] 0;



     foreach ($computer in $ComputerName) {
        if ($Credentials) {
           $os=Get-WMIObject win32_operatingsystem -computername $computer -Credential $Credentials
        } else {
           $os=Get-WMIObject win32_operatingsystem -computername $computer
        }

        $osArchitecture = $os.OSArchitecture

        if ($Credentials) {
           $regProv = Get-Wmiobject -list "StdRegProv" -namespace root\default -computername $computer -Credential $Credentials
        } else {
           $regProv = Get-Wmiobject -list "StdRegProv" -namespace root\default -computername $computer
        }

        $VersionList = New-Object -TypeName System.Collections.ArrayList
        $PathList = New-Object -TypeName System.Collections.ArrayList
        $PackageList = New-Object -TypeName System.Collections.ArrayList
        $ClickToRunPathList = New-Object -TypeName System.Collections.ArrayList
        $ConfigItemList = New-Object -TypeName System.Collections.ArrayList
        $ClickToRunList = new-object PSObject[] 0;

        foreach ($regKey in $officeKeys) {
           $officeVersion = $regProv.EnumKey($HKLM, $regKey)
           foreach ($key in $officeVersion.sNames) {
              if ($key -match "\d{2}\.\d") {
                if (!$VersionList.Contains($key)) {
                  $AddItem = $VersionList.Add($key)
                }

                $path = join-path $regKey $key

                $configPath = join-path $path "Common\Config"
                $configItems = $regProv.EnumKey($HKLM, $configPath)
                foreach ($configId in $configItems.sNames) {
                   $Add = $ConfigItemList.Add($configId.ToUpper())
                }

                $cltr = New-Object -TypeName PSObject
                $cltr | Add-Member -MemberType NoteProperty -Name InstallPath -Value ""
                $cltr | Add-Member -MemberType NoteProperty -Name UpdatesEnabled -Value $false
                $cltr | Add-Member -MemberType NoteProperty -Name UpdateUrl -Value ""
                $cltr | Add-Member -MemberType NoteProperty -Name StreamingFinished -Value $false
                $cltr | Add-Member -MemberType NoteProperty -Name Platform -Value ""
                $cltr | Add-Member -MemberType NoteProperty -Name ClientCulture -Value ""
            
                $packagePath = join-path $path "Common\InstalledPackages"
                $clickToRunPath = join-path $path "ClickToRun\Configuration"
                $virtualInstallPath = $regProv.GetStringValue($HKLM, $clickToRunPath, "InstallationPath").sValue

                if ($virtualInstallPath) {

                } else {
                  $clickToRunPath = join-path $regKey "ClickToRun\Configuration"
                  $virtualInstallPath = $regProv.GetStringValue($HKLM, $clickToRunPath, "InstallationPath").sValue
                }

                if ($virtualInstallPath) {
                   if (!$ClickToRunPathList.Contains($virtualInstallPath.ToUpper())) {
                      $AddItem = $ClickToRunPathList.Add($virtualInstallPath.ToUpper())
                   }

                   $cltr.InstallPath = $virtualInstallPath
                   $cltr.StreamingFinished = $regProv.GetStringValue($HKLM, $clickToRunPath, "StreamingFinished").sValue
                   $cltr.UpdatesEnabled = $regProv.GetStringValue($HKLM, $clickToRunPath, "UpdatesEnabled").sValue
                   $cltr.UpdateUrl = $regProv.GetStringValue($HKLM, $clickToRunPath, "UpdateUrl").sValue
                   $cltr.Platform = $regProv.GetStringValue($HKLM, $clickToRunPath, "Platform").sValue
                   $cltr.ClientCulture = $regProv.GetStringValue($HKLM, $clickToRunPath, "ClientCulture").sValue
                   $ClickToRunList += $cltr
                }

                $packageItems = $regProv.EnumKey($HKLM, $packagePath)
                $officeItems = $regProv.EnumKey($HKLM, $path)

                foreach ($itemKey in $officeItems.sNames) {
                  $itemPath = join-path $path $itemKey
                  $installRootPath = join-path $itemPath "InstallRoot"

                  $filePath = $regProv.GetStringValue($HKLM, $installRootPath, "Path").sValue
                  if (!$PathList.Contains($filePath)) {
                      $AddItem = $PathList.Add($filePath)
                  }
                }

                foreach ($packageGuid in $packageItems.sNames) {
                  $packageItemPath = join-path $packagePath $packageGuid
                  $packageName = $regProv.GetStringValue($HKLM, $packageItemPath, "").sValue
            
                  if (!$PackageList.Contains($packageName)) {
                    $AddItem = $PackageList.Add($packageName.Replace(' ', '').ToLower())
                  }
                }

              }
           }
        }

    

        foreach ($regKey in $installKeys) {
            $keyList = new-object System.Collections.ArrayList
            $keys = $regProv.EnumKey($HKLM, $regKey)

            foreach ($key in $keys.sNames) {
               $path = join-path $regKey $key
               $installPath = $regProv.GetStringValue($HKLM, $path, "InstallLocation").sValue
               if ($installPath.Length -eq 0) { continue }

               $buildType = "64-Bit"
               if ($osArchitecture -eq "32-bit") {
                  $buildType = "32-Bit"
               }

               if ($regKey.ToUpper().Contains("Wow6432Node".ToUpper())) {
                  $buildType = "32-Bit"
               }

               if ($key -match "{.{8}-.{4}-.{4}-1000-0000000FF1CE}") {
                  $buildType = "64-Bit" 
               }

               if ($key -match "{.{8}-.{4}-.{4}-0000-0000000FF1CE}") {
                  $buildType = "32-Bit" 
               }

               if ($modifyPath) {
                   if ($modifyPath.ToLower().Contains("platform=x86")) {
                      $buildType = "32-Bit"
                   }

                   if ($modifyPath.ToLower().Contains("platform=x64")) {
                      $buildType = "64-Bit"
                   }
               }

               $primaryOfficeProduct = $false
               $officeProduct = $false
               foreach ($officeInstallPath in $PathList) {
                 if ($officeInstallPath) {
                    $installReg = "^" + $installPath.Replace('\', '\\')
                    $installReg = $installReg.Replace('(', '\(')
                    $installReg = $installReg.Replace(')', '\)')
                    if ($officeInstallPath -match $installReg) { $officeProduct = $true }
                 }
               }

               if (!$officeProduct) { continue };
           
               $name = $regProv.GetStringValue($HKLM, $path, "DisplayName").sValue          

               if ($ConfigItemList.Contains($key.ToUpper()) -and $name.ToUpper().Contains("MICROSOFT OFFICE")) {
                  $primaryOfficeProduct = $true
               }

               $version = $regProv.GetStringValue($HKLM, $path, "DisplayVersion").sValue
               $modifyPath = $regProv.GetStringValue($HKLM, $path, "ModifyPath").sValue 

               $cltrUpdatedEnabled = $NULL
               $cltrUpdateUrl = $NULL
               $clientCulture = $NULL;

               [string]$clickToRun = $false
               if ($ClickToRunPathList.Contains($installPath.ToUpper())) {
                   $clickToRun = $true
                   if ($name.ToUpper().Contains("MICROSOFT OFFICE")) {
                      $primaryOfficeProduct = $true
                   }

                   foreach ($cltr in $ClickToRunList) {
                     if ($cltr.InstallPath) {
                       if ($cltr.InstallPath.ToUpper() -eq $installPath.ToUpper()) {
                           $cltrUpdatedEnabled = $cltr.UpdatesEnabled
                           $cltrUpdateUrl = $cltr.UpdateUrl
                           if ($cltr.Platform -eq 'x64') {
                               $buildType = "64-Bit" 
                           }
                           if ($cltr.Platform -eq 'x86') {
                               $buildType = "32-Bit" 
                           }
                           $clientCulture = $cltr.ClientCulture
                       }
                     }
                   }
               }
           
               if (!$primaryOfficeProduct) {
                  if (!$ShowAllInstalledProducts) {
                      continue
                  }
               }

           

               $object = New-Object PSObject -Property @{ComputerName = $computer; DisplayName = $name; Version = $version; ClickToRun = $clickToRun }
               $object | Add-Member MemberSet PSStandardMembers $PSStandardMembers
               $results += $object
  
               
            }

            
         }

         return $results

         
      }

      
     
    }

}

function Get-ScheduledTasks{
    [cmdletbinding()]                        
    param (                        
        [parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]                        
        [string[]] $ComputerName = $env:computername,                        
        [string] $TaskName                        
    )    
                        
    foreach ($Computer in $ComputerName) {                        
        $SchService = New-Object -ComObject Schedule.Service                        
        $SchService.Connect($Computer)                        
        $Rootfolder = $SchService.GetFolder("\")            
        $folders = @($RootFolder)             
        $folders += Get-Tasksubfolders -FolderRef $RootFolder
                                
            foreach($Folder in $folders) {                        
                $Tasks = $folder.gettasks(1)                        
                foreach($Task in $Tasks) {                        
                    $OutputObj = New-Object -TypeName PSobject                         
                    $OutputObj | Add-Member -MemberType NoteProperty -Name ComputerName -Value $Computer                        
                    $OutputObj | Add-Member -MemberType NoteProperty -Name TaskName -Value $Task.Name                        
                    $OutputObj | Add-Member -MemberType NoteProperty -Name TaskFolder -Value $Folder.path                        
                    $OutputObj | Add-Member -MemberType NoteProperty -Name IsEnabled -Value $task.enabled                        
                    $OutputObj | Add-Member -MemberType NoteProperty -Name LastRunTime -Value $task.LastRunTime                        
                    $OutputObj | Add-Member -MemberType NoteProperty -Name NextRunTime -Value $task.NextRunTime                        
                    if($TaskName) {                        
                        if($Task.Name -eq $TaskName) {                        
                            $OutputObj                        
                        }                        
                    } 
                    else{                        
                        $OutputObj                        
                    }                         
                }                        
            }                        
    }
}

function Nuke-Office([string[]] $ComputerName){

    foreach($computer in $ComputerName){
       
        $osVersion = [System.Environment]::OSVersion.Version | foreach {"$_.Major"}
        $versionTest = Get-OfficeVersion $computer
        $c2r = $versionTest.ClicktoRun
        $CurrentDate = Get-Date
        $CurrentDate = $CurrentDate.ToString('MM-dd-yyyy hh:mm:ss')
        $destination = "\\$computer\c$\Windows\Temp"
        $log = $computer + "Log.txt"
        $logPath = "\\" + $computer + "\c$\Windows\Temp\$log"
        
            
        if($c2r -eq $true){

            $taskName = Get-ScheduledTasks $computer | foreach {$_.TaskName}

            Copy-Item -Path ".\Offscrubc2r.vbs" -Destination $destination -Force | Out-Null

                if(Test-Path -Path $destination){

                    "$CurrentDate - OffScrubc2r.vbs has been copied to C:\Windows\Temp" | Out-File $logPath
                }
                else{

                    "$CurrentDate - Unable to copy OffScrubc2r.vbs to $destination" | Out-File $logPath
                }
                    
            if(!($taskName -match "OffScrub")){

                [string]$TaskRun = '"%systemroot%\system32\cscript.exe C:\Windows\Temp\OffScrubc2r.vbs"'
                $command = "schtasks.exe /create /s $computer /ru System /tn OffScrub /tr $TaskRun /sc Onstart"
                $runTask = "schtasks.exe /run /s $computer /tn OffScrub"
                
                ac $logPath "$CurrentDate - Attempting to create the scheduled task on $computer..."

                Invoke-Expression $command >> $logPath

                ac $logPath "$CurrentDate - Attempting to run the new task on $computer..."

                Invoke-Expression $runTask >> $logPath

            }
            else{

                [string]$TaskRun = '"%systemroot%\system32\cscript.exe C:\Windows\Temp\OffScrubc2r.vbs"'
                $command = "schtasks.exe /create /s $computer /ru System /tn OffScrub /tr $TaskRun /sc Once /sd 01/01/2999 /st 00:00"
                $runTask = "schtasks.exe /run /s $computer /tn OffScrub"

                ac $logPath "$CurrentDate - The scheduled task already exists on $computer. Attempting to run the task..."

                Invoke-Expression $runTask >> $logPath
            }
        }
        else{
            if($versionTest.Version -match "11.*"){

                if($osVersion -match "5.1.*" -or $osVersion -match "5.2.*" -or $osVersion -match "6.0.*" -or $osVersion -match "6.1.*"){

                $taskName = Get-ScheduledTasks $computer | foreach {$_.TaskName}

                    if(!($taskName -match "OffScrub")){

                        [string]$TaskRun = '"%systemroot%\system32\cscript.exe C:\Windows\Temp\MicrosoftFixit50416.msi"'
                        $command = "schtasks.exe /create /s $computer /ru System /tn OffScrub /tr $TaskRun /sc Once /sd 01/01/2999 /st 00:00"
                        $runTask = "schtasks.exe /run /s $computer /tn OffScrub"
                
                        Copy-Item -Path ".\MicrosoftFixit50416.msi" -Destination "\\$computer\c$\Windows\Temp" -Force | Out-Null

                        ac $logPath "$CurrentDate - Attempting to create the scheduled task on $computer..."

                        Invoke-Expression $command >> $logPath

                        ac $logPath "$CurrentDate - Attempting to run the new task on $computer..."

                        Invoke-Expression $runTask >> $logPath
                    }
                    else{

                        [string]$TaskRun = '"%systemroot%\system32\cscript.exe C:\Windows\Temp\MicrosoftFixit50416.msi"'
                        $command = "schtasks.exe /create /s $computer /ru System /tn OffScrub /tr $TaskRun /sc Once /sd 01/01/2999 /st 00:00"
                        $runTask = "schtasks.exe /run /s $computer /tn OffScrub"

                        ac $logPath "$CurrentDate - The scheduled task already exists on $computer. Attempting to run the task..."

                        Invoke-Expression $runTask >> $logPath
                    }
                }
                else{

                    $taskName = Get-ScheduledTasks $computer | foreach {$_.TaskName}

                    if(!($taskName -match "OffScrub")){

                        [string]$TaskRun = '"%systemroot%\system32\cscript.exe C:\Windows\Temp\OffScrub03.vbs"'
                        $command = "schtasks.exe /create /s $computer /ru System /tn OffScrub /tr $TaskRun /sc Once /sd 01/01/2999 /st 00:00"
                        $runTask = "schtasks.exe /run /s $computer /tn OffScrub"
                
                        Copy-Item -Path ".\Offscrub03.vbs" -Destination "\\$computer\c$\Windows\Temp" -Force | Out-Null

                        ac $logPath "$CurrentDate - Attempting to create the scheduled task on $computer..."

                        Invoke-Expression $command >> $logPath

                        ac $logPath "$CurrentDate - Attempting to run the new task on $computer..."

                        Invoke-Expression $runTask >> $logPath
                    }
                    else{

                        [string]$TaskRun = '"%systemroot%\system32\cscript.exe C:\Windows\Temp\OffScrub03.vbs"'
                        $command = "schtasks.exe /create /s $computer /ru System /tn OffScrub /tr $TaskRun /sc Once /sd 01/01/2999 /st 00:00"
                        $runTask = "schtasks.exe /run /s $computer /tn OffScrub"

                        ac $logPath "$CurrentDate - The scheduled task already exists on $computer."

                        Invoke-Expression $runTask >> $logPath
                    }
                }
            }
            elseif($versionTest.Version -match "12.*"){

                if($osVersion -match "5.1.*" -or $osVersion -match "5.2.*" -or $osVersion -match "6.0.*" -or $osVersion -match "6.1.*"){

                $taskName = Get-ScheduledTasks $computer | foreach {$_.TaskName}

                    if(!($taskName -match "OffScrub")){

                        [string]$TaskRun = '"%systemroot%\system32\cscript.exe C:\Windows\Temp\MicrosoftFixit50154.msi"'
                        $command = "schtasks.exe /create /s $computer /ru System /tn OffScrub /tr $TaskRun /MO ONSTART"
                        $runTask = "schtasks.exe /run /s $computer /tn OffScrub"
                
                        Copy-Item -Path ".\MicrosoftFixit50154.msi" -Destination "\\$computer\c$\Windows\Temp" -Force | Out-Null

                        ac $logPath "$CurrentDate - Attempting to create the scheduled task on $computer..."

                        Invoke-Expression $command >> $logPath

                        ac $logPath "$CurrentDate - Attempting to run the new task on $computer..."

                        Invoke-Expression $runTask >> $logPath
                    }
                    else{

                        [string]$TaskRun = '"%systemroot%\system32\cscript.exe C:\Windows\Temp\MicrosoftFixit50154.msi"'
                        $command = "schtasks.exe /create /s $computer /ru System /tn OffScrub /tr $TaskRun /MO ONSTART"
                        $runTask = "schtasks.exe /run /s $computer /tn OffScrub"

                        ac $logPath "$CurrentDate - The scheduled task already exists on $computer..."

                        Invoke-Expression $runTask >> $logPath
                    }
                }
                else{

                    $taskName = Get-ScheduledTasks $computer | foreach {$_.TaskName}

                    if(!($taskName -match "OffScrub")){

                        [string]$TaskRun = '"%systemroot%\system32\cscript.exe C:\Windows\Temp\OffScrub07.vbs"'
                        $command = "schtasks.exe /create /s $computer /ru System /tn OffScrub /tr $TaskRun /sc Once /sd 01/01/2999 /st 00:00"
                        $runTask = "schtasks.exe /run /s $computer /tn OffScrub"
                
                        Copy-Item -Path ".\Offscrub07.vbs" -Destination "\\$computer\c$\Windows\Temp" -Force | Out-Null

                        ac $logPath "$CurrentDate - Attempting to create the scheduled task on $computer..."

                        Invoke-Expression $command >> $logPath

                        ac $logPath "$CurrentDate - Attempting to run the new task on $computer..."

                        Invoke-Expression $runTask >> $logPath
                    }
                    else{

                        [string]$TaskRun = '"%systemroot%\system32\cscript.exe C:\Windows\Temp\OffScrub07.vbs"'
                        $command = "schtasks.exe /create /s $computer /ru System /tn OffScrub /tr $TaskRun /sc Once /sd 01/01/2999 /st 00:00"
                        $runTask = "schtasks.exe /run /s $computer /tn OffScrub"

                        ac $logPath "$CurrentDate - The scheduled task already exists on $computer."

                        Invoke-Expression $runTask >> $logPath
                    }
                }
            }
            elseif($versionTest.Version -match "14.*"){

                if($osVersion -match "5.1.*" -or $osVersion -match "5.2.*" -or $osVersion -match "6.0.*" -or $osVersion -match "6.1.*"){

                $taskName = Get-ScheduledTasks $computer | foreach {$_.TaskName}

                    if(!($taskName -match "OffScrub")){

                        [string]$TaskRun = '"%systemroot%\system32\cscript.exe C:\Windows\Temp\OffScrub10.vbs"'
                        $command = "schtasks.exe /create /s $computer /ru System /tn OffScrub /tr $TaskRun /sc Once /sd 01/01/2999 /st 00:00"
                        $runTask = "schtasks.exe /run /s $computer /tn OffScrub"
                
                        Copy-Item -Path ".\OffScrub10.vbs" -Destination "\\$computer\c$\Windows\Temp" -Force | Out-Null

                        ac $logPath "$CurrentDate - Attempting to create the scheduled task on $computer..."

                        Invoke-Expression $command >> $logPath

                        ac $logPath "$CurrentDate - Attempting to run the new task on $computer..."

                        Invoke-Expression $runTask >> $logPath
                    }
                    else{

                        [string]$TaskRun = '"%systemroot%\system32\cscript.exe C:\Windows\Temp\OffScrub10.vbs"'
                        $command = "schtasks.exe /create /s $computer /ru System /tn OffScrub /tr $TaskRun /sc Once /sd 01/01/2999 /st 00:00"
                        $runTask = "schtasks.exe /run /s $computer /tn OffScrub"

                        ac $logPath "$CurrentDate - The scheduled task already exists on $computer."

                        Invoke-Expression $runTask >> $logPath
                    }
                }
                 else{

                    $taskName = Get-ScheduledTasks $computer | foreach {$_.TaskName}

                    if(!($taskName -match "OffScrub")){

                        [string]$TaskRun = '"%systemroot%\system32\cscript.exe C:\Windows\Temp\OffScrub10.vbs"'
                        $command = "schtasks.exe /create /s $computer /ru System /tn OffScrub /tr $TaskRun /sc Once /sd 01/01/2999 /st 00:00"
                        $runTask = "schtasks.exe /run /s $computer /tn OffScrub"
                
                        Copy-Item -Path ".\Offscrub10.vbs" -Destination "\\$computer\c$\Windows\Temp" -Force | Out-Null

                        ac $logPath "$CurrentDate - Attempting to create the scheduled task on $computer..."

                        Invoke-Expression $command >> $logPath

                        ac $logPath "$CurrentDate - Attempting to run the new task on $computer..."

                        Invoke-Expression $runTask >> $logPath
                    }
                    else{

                        [string]$TaskRun = '"%systemroot%\system32\cscript.exe C:\Windows\Temp\OffScrub10.vbs"'
                        $command = "schtasks.exe /create /s $computer /ru System /tn OffScrub /tr $TaskRun /sc Once /sd 01/01/2999 /st 00:00"
                        $runTask = "schtasks.exe /run /s $computer /tn OffScrub"

                        ac $logPath "$CurrentDate - The scheduled task already exists on $computer."

                        Invoke-Expression $runTask >> $logPath
                    }
                }
            }
            elseif($versionTest.Version -match "15.*"){

                $taskName = Get-ScheduledTasks $computer | foreach {$_.TaskName}

                    if(!($taskName -match "OffScrub")){

                        [string]$TaskRun = '"%systemroot%\system32\cscript.exe C:\Windows\Temp\OffScrub_O15msi.vbs"'
                        $command = "schtasks.exe /create /s $computer /ru System /tn OffScrub /tr $TaskRun /sc Once /sd 01/01/2999 /st 00:00"
                        $runTask = "schtasks.exe /run /s $computer /tn OffScrub"
                
                        Copy-Item -Path ".\OffScrub_O15msi.vbs" -Destination "\\$computer\c$\Windows\Temp" -Force | Out-Null

                        ac $logPath "$CurrentDate - Attempting to create the scheduled task on $computer..."

                        Invoke-Expression $command >> $logPath

                        ac $logPath "$CurrentDate - Attempting to run the new task on $computer..."

                        Invoke-Expression $runTask >> $logPath
                    }
                    else{

                        [string]$TaskRun = '"%systemroot%\system32\cscript.exe C:\Windows\Temp\OffScrub_O15msi.vbs"'
                        $command = "schtasks.exe /create /s $computer /ru System /tn OffScrub /tr $TaskRun /sc Once /sd 01/01/2999 /st 00:00"
                        $runTask = "schtasks.exe /run /s $computer /tn OffScrub"

                        ac $logPath "$CurrentDate - The scheduled task already exists on $computer."

                        Invoke-Expression $runTask >> $logPath
                    }
            }
            elseif($versionTest.Version -match "16.*"){

                $taskName = Get-ScheduledTasks $computer | foreach {$_.TaskName}

                    if(!($taskName -match "OffScrub")){

                        [string]$TaskRun = '"%systemroot%\system32\cscript.exe C:\Windows\Temp\OffScrub_O16msi.vbs"'
                        $command = "schtasks.exe /create /s $computer /ru System /tn OffScrub /tr $TaskRun /sc Once /sd 01/01/2999 /st 00:00"
                        $runTask = "schtasks.exe /run /s $computer /tn OffScrub"
                
                        Copy-Item -Path ".\OffScrub_O16msi.vbs" -Destination "\\$computer\c$\Windows\Temp" -Force | Out-Null

                        ac $logPath "$CurrentDate - Attempting to create the scheduled task on $computer..."

                        Invoke-Expression $command >> $logPath

                        ac $logPath "$CurrentDate - Attempting to run the new task on $computer..."

                        Invoke-Expression $runTask >> $logPath
                    }
                    else{

                        [string]$TaskRun = '"%systemroot%\system32\cscript.exe C:\Windows\Temp\OffScrub_O16msi.vbs"'
                        $command = "schtasks.exe /create /s $computer /ru System /tn OffScrub /tr $TaskRun /sc Once /sd 01/01/2999 /st 00:00"
                        $runTask = "schtasks.exe /run /s $computer /tn OffScrub"

                        ac $logPath "$CurrentDate - The scheduled task already exists on $computer."

                        Invoke-Expression $runTask >> $logPath
                    }
            }   
        }
    }
}

}