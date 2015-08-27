function Find-AddinInformation {

    # Find a list of available add-ins
    function Find-ComAddins {

        Write-Host "Looking for COM add-ins..."`n

        $addinOfficePath = "HKCU:\Software\Microsoft\Office"

        Get-ChildItem -Path $addinOfficePath -Recurse | Where-Object { $_.PsChildName -match 'Addins' } | Get-ChildItem | Get-ItemProperty | select PSChildName,FriendlyName | Where-Object {$_ -ne ""}

    }

    # Look for the uninstall strings
    function Find-UninstallLocations {

        Write-Host "Looking for installed applications..."`n

        [Management.ManagementBaseObject] $os = Get-WMIObject win32_operatingsystem
        if ($os.OSArchitecture -match "^64")
        {
            $bit64 = "Wow6432Node"
        }
    
        $uninstallLocation = Get-ChildItem Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall -Recurse | Get-ItemProperty | Select DisplayName, UninstallString

        $uninstallLocation64 = Get-ChildItem Registry::HKEY_LOCAL_MACHINE\SOFTWARE\$bit64\Microsoft\Windows\CurrentVersion\Uninstall -Recurse | Get-ItemProperty | Select DisplayName, UninstallString

        $allUninstallLocations = $uninstallLocation,$uninstallLocation64
    
        $allUninstallLocations | Out-File -FilePath "$env:TEMP\Uninstalls.txt" 

        Get-Content "$env:TEMP\Uninstalls.txt" | ? {$_.trim() -ne ""} | Set-Content "$env:TEMP\Uninstalls.txt"

        Get-Content "$env:TEMP\Uninstalls.txt"

}

    # Look for Excel add-in files (xla,xlam)
    function Find-AddinExt {

        Write-Host "Looking for add-in files by extension..."`n

        
        $dir86 = Get-ChildItem ${env:ProgramFiles(x86)} -Recurse
        $dir = Get-ChildItem $env:ProgramFiles -Recurse
        $list = $dir | where { $_.Extension -eq ".xla" -or $_.Extension -eq ".xlam" -or $_.Extension -eq ".xll" -or $_.Extension -eq ".ppa" -or $_.Extension -eq ".ppam" -or $_.Extension -eq ".pa" -or $_.Extension -eq ".accda"  -or $_.Extension -eq ".mda" -or $_.Extension -eq ".wll" }
        $list86 = $dir86 | where { $_.Extension -eq ".xla" -or $_.Extension -eq ".xlam" -or $_.Extension -eq ".xll" -or $_.Extension -eq ".ppa" -or $_.Extension -eq ".ppam" -or $_.Extension -eq ".pa" -or $_.Extension -eq ".accda"  -or $_.Extension -eq ".mda" -or $_.Extension -eq ".wll" }
        $files = $list,$list86 | foreach {$_.Name} | Out-Null

         
        if(Test-Path "$env:TEMP\Uninstalls.txt")
        {

            Get-Content "$env:TEMP\Uninstalls.txt" | foreach {$_.DisplayName -match "Favorite"}
                
            return $true
        }
        
    }
   
        
    

    # Look in the AddinLoadTimes for enabled add-ins
    function Find-AddInLoadTime {

        Write-Host "Looking for enabled add-ins..."`n
        
        Get-ChildItem -Path "Registry::HKEY_USERS\S-1-5-21-*\Software\Microsoft\Office" -Recurse | Where-Object { $_.PsChildName -match 'AddInLoadTimes' } | foreach {$_.Property}

    }


    
    Find-UninstallLocations
    
    Find-ApplicationAddIns

    Find-AddinExt
    
    Find-AddInLoadTime

}