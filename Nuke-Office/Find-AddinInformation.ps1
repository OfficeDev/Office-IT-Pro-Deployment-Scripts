function Remove-Addins {

    # Find a list of available add-ins
    function Find-ComAddins {

        Write-Host "Looking for COM add-ins..."`n

        $addinOfficePath = "HKCU:\Software\Microsoft\Office"

        $result = Get-ChildItem -Path $addinOfficePath -Recurse | Where-Object { $_.PsChildName -match 'Addins' } | Get-ChildItem | Get-ItemProperty | select PSChildName,FriendlyName

        $result
    }

    # Look for the uninstall strings
    function Find-UninstallLocations {

        Write-Host "Looking for installed applications..."`n

        [Management.ManagementBaseObject] $os = Get-WMIObject win32_operatingsystem
        if ($os.OSArchitecture -match "^64")
        {
            $bit64 = "Wow6432Node"
        }
    
        $uninstallLocation = Get-ChildItem Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall -Recurse | 
        Get-ItemProperty | Select DisplayName, UninstallString, QuietUninstallString

        $uninstallLocation64 = Get-ChildItem Registry::HKEY_LOCAL_MACHINE\SOFTWARE\$bit64\Microsoft\Windows\CurrentVersion\Uninstall -Recurse | 
        Get-ItemProperty | Select DisplayName, UninstallString, QuietUninstallString

        $allUninstallLocations = $uninstallLocation,$uninstallLocation64

        $allUninstallLocations
    }

    # Look for Excel add-in files (xla,xlam, etc)
    function Find-AddinExt {

        Write-Host "Looking for add-in files by extension..."`n

        
        $dir86 = Get-ChildItem ${env:ProgramFiles(x86)} -Recurse
        $dir = Get-ChildItem $env:ProgramFiles -Recurse
        $list = $dir | where { $_.Extension -eq ".xla" -or $_.Extension -eq ".xlam" -or $_.Extension -eq ".xll" -or $_.Extension -eq ".ppa" -or $_.Extension -eq ".ppam" -or $_.Extension -eq ".pa" -or $_.Extension -eq ".accda"  -or $_.Extension -eq ".mda" -or $_.Extension -eq ".wll" }
        $list86 = $dir86 | where { $_.Extension -eq ".xla" -or $_.Extension -eq ".xlam" -or $_.Extension -eq ".xll" -or $_.Extension -eq ".ppa" -or $_.Extension -eq ".ppam" -or $_.Extension -eq ".pa" -or $_.Extension -eq ".accda"  -or $_.Extension -eq ".mda" -or $_.Extension -eq ".wll" }
        $files = $list,$list86 | foreach {$_.BaseName}
        $files
    }
    
    $UninstallInfo = Find-UninstallLocations
    
    $ComAddins = Find-ComAddins

    $ExtAddins = Find-AddinExt

    $Addins = $ComAddins.FriendlyName
    $Addins += $ComAddins.PSChildName
    $Addins += $ExtAddins

    foreach($addin in $Addins){
        foreach($Uninstall in $UninstallInfo){
            
            if($Uninstall -eq $addin){
                if($Uninstall.QuietUninstallString -ne $null){
                    Invoke-Expression $Uninstall.QuietUninstallString | Out-Null
                }else{
                    if($Uninstall.UninstallString -match "*MsiExec.exe*"){
                        Invoke-Expression "$($Uninstall.UninstallString) /q" | Out-Null
                    }elseif($Uninstall.UninstallString -match "*.exe*"){
                        Invoke-Expression "$($Uninstall.UninstallString) /SILENT" | Out-Null
                    }
                }
                break;
            }
        }
    }

}