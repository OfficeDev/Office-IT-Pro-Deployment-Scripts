function Get-OfficeAddins {
Param(
    [string]$ComputerName = $env:COMPUTERNAME
)

    $defaultDisplaySet = 'ComputerName','Application','Name','Description','FriendlyName','LoadBehavior','RegistryPath','FullPath','LoadTime','OfficeVersion','Type'
    $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet',[string[]]$defaultDisplaySet)
    $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
    $results = New-Object PSObject[] 0;
    
    $HKCU = [UInt32] "0x80000001"
    $HKLM = [UInt32] "0x80000002"
    $HKU = [UInt32] "0x80000003"
    
    $HKEYS = @($HKCU, $HKLM)
    
    $officeApps = @("Word","Excel","PowerPoint","Outlook","MS Project")
    
    $COMKeys = @("SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office",
                 "SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Visio\Addins",
                 "SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\USERS\.DEFAULT\Software\Microsoft\Office",
                 "SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\USERS\.DEFAULT\Software\Microsoft\Visio\Addins")
    
    $VSTOKeys = @("Software\Microsoft\Office",  
                  "Software\Wow6432Node\Microsoft\Office",
                  "Software\Microsoft",
                  "Software\Wow6432Node")
    
    $regProv = Get-WmiObject -List "StdRegProv" -Namespace root\default -ComputerName $ComputerName
    
    foreach($comkey in $COMKeys){
        if($comkey -notmatch "Office"){
            $searchApps = "Visio"
        } else {
            $searchApps = $officeApps
        }

        foreach($officeapp in $searchApps){
            $path = Join-Path $comkey $officeapp
    
            foreach($hkey in $hkeys){
                $hkeyEnum = $regProv.EnumKey($hkey, $path)
    
                if($hkeyEnum.sNames -contains "Addins"){
                    $addinsPath = Join-Path $path "Addins"
                    $addinEnum = $regProv.EnumKey($hkey, $addinsPath)
                    foreach($addinapp in $addinEnum.sNames){
                        $addinpath = Join-Path $addinsPath $addinapp
                        
                        switch($hkey){
                            '2147483649'{
                                $hive = 'HKCU'
                            }
                            '2147483650'{
                                $hive = 'HKLM'
                            }
                        }
                        
                        $LoadBehavior = ($regprov.GetDWORDValue($hkey, $addinPath, 'LoadBehavior')).uValue
                        $Description = ($regProv.GetStringValue($hkey, $addinPath, 'Description')).sValue
                        $FriendlyName = ($regProv.GetStringValue($hkey, $addinPath, 'FriendlyName')).sValue
                        $FullPath = Get-AddinFullPath -AddinID $addinapp -AddinType "COM"
                        $loadTime = Get-AddinLoadtime -AddinID $addinapp
                        $addinOfficeVersion = Get-AddinOfficeVersion -AddinID $addinapp
    
                        $object = New-Object PSObject -Property @{ComputerName = $ComputerName; Application = $officeapp; Name = $addinapp; RegistryPath = $addinpath; Type = "COM"; 
                                                                  Description = $Description; FriendlyName = $FriendlyName; LoadBehavior = $LoadBehavior;
                                                                  FullPath = $FullPath; LoadTime = $loadTime; OfficeVersion = $addinOfficeVersion}
                        $object | Add-Member MemberSet PSStandardMembers $PSStandardMembers
                        $results += $object
                    }
                }
            }
        }
    }

    foreach($vstokey in $VSTOKeys){
        if($vstokey -notmatch "Office"){
            $searchApps = "Visio"
        } else {
            $searchApps = $officeApps
        }
        foreach($officeapp in $searchApps){
            $path = Join-Path $vstokey $officeapp
        
            foreach($hkey in $hkeys){
                $hkeyEnum = $regProv.EnumKey($hkey, $path)
        
                if($hkeyEnum.sNames -contains "Addins"){
                    $addinsPath = Join-Path $path "Addins"
                    $addinEnum = $regProv.EnumKey($hkey, $addinsPath)
                    foreach($addinapp in $addinEnum.sNames){
                        $addinpath = Join-Path $addinsPath $addinapp
        
                        switch($hkey){
                            '2147483649'{
                                $hive = 'HKCU'
                            }
                            '2147483650'{
                                $hive = 'HKLM'
                            }
                        }
                      
                        $LoadBehavior = ($regprov.GetDWORDValue($hkey, $addinPath, 'LoadBehavior')).uValue
                        $Description = ($regProv.GetStringValue($hkey, $addinPath, 'Description')).sValue
                        $FriendlyName = ($regProv.GetStringValue($hkey, $addinPath, 'FriendlyName')).sValue
                        $FullPath = Get-AddinFullPath -AddinID $addinapp -AddinType "VSTO"
                        $loadTime = Get-AddinLoadtime -AddinID $addinapp
                        $addinOfficeVersion = Get-AddinOfficeVersion -AddinID $addinapp
        
                        $object = New-Object PSObject -Property @{ComputerName = $ComputerName; Application = $officeapp; Name = $addinapp; RegistryPath = $addinpath; Type = "VSTO"; 
                                                                  Description = $Description; FriendlyName = $FriendlyName; LoadBehavior = $LoadBehavior;
                                                                  FullPath = $FullPath; LoadTime = $loadTime; OfficeVersion = $addinOfficeVersion}
                        $object | Add-Member MemberSet PSStandardMembers $PSStandardMembers
                        $results += $object
                    }
                }
            }
        }
    }
    
    return $results;

}

function Get-AddinFullPath {
Param(
    [string]$ComputerName = $env:COMPUTERNAME,
    [string]$AddinID,
    [string]$AddinType
)

    $regProv = Get-WmiObject -List "StdRegProv" -Namespace root\default -ComputerName $ComputerName

    $HKLM = [UInt32] "0x80000002"

    $comPathKeys = @("SOFTWARE\Classes\CLSID",
                     "SOFTWARE\Classes\Wow6432Node\CLSID",
                     "SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\SOFTWARE\Classes\CLSID",
                     "SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\SOFTWARE\Classes\Wow6432Node\CLSID")

    switch($AddinType){
        "COM"{
            $clsid = Get-CLSID -ProgId $AddinID
            foreach($key in $comPathKeys){
                $path = Join-Path $key $clsid
                $InProcPath = Join-Path $path "InprocServer32"
                if(Test-Path "HKLM:\$InProcPath"){
                    $fullpath = Get-ItemProperty ("HKLM:\$InProcPath")
                    $fullpath = $fullpath.'(default)'
                }
            }
        }
        "VSTO"{
            $fullpath = Get-ManifestKey -AddinID $AddinID
        }

    }

    return $fullpath;

}

function Get-CLSID {
Param(
    [string]$ComputerName = $env:COMPUTERNAME,
    [string]$ProgId
)
    $regProv = Get-WmiObject -List "StdRegProv" -Namespace root\default -ComputerName $ComputerName

    $HKLM = [UInt32] "0x80000002"

    $ClsIdPaths = @("SOFTWARE\Classes\Wow6432Node\CLSID",
                    "SOFTWARE\Classes\CLSID",
                    "SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\SOFTWARE\Classes\Wow6432Node\CLSID",
                    "SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\SOFTWARE\Classes\CLSID")

    foreach($ClsIdPath in $ClsIdPaths){
        $Clsids = $regProv.EnumKey($HKLM, $ClsIdPath)
        $clsids = $Clsids.sNames

        foreach($clsid in $Clsids){
            if($Clsid -match "{.{8}-.{4}-.{4}-.{4}-.{12}}"){
                $path = Join-Path $ClsIdPath $clsid
                $progIdPath = Join-Path $path "ProgID"
                $literalPath = "HKLM:\" + $path

                $ProgIDValue = Get-ChildItem $literalPath | ForEach-Object {
                    if($_.PSChildName -eq "ProgID"){
                        $_.GetValue("")
                    }
                }
                
                if($ProgIDValue -match $ProgId){
                    $InprocServer32 = Get-ChildItem $literalPath | ForEach-Object {
                        if($_.PSChildName -eq "InprocServer32"){
                            $_.GetValue("")
                        }
                    }

                    return $clsid
                }
            }
        }
    }
}

function Get-ManifestKey {
Param(
    [string]$ComputerName = $env:COMPUTERNAME,
    [string]$AddinID
)

    $regProv = Get-WmiObject -List "StdRegProv" -Namespace root\default -ComputerName $ComputerName

    $HKLM = [UInt32] "0x80000002"
    $HKU = [UInt32] "0x80000003"

    $hkeys = @($HKLM,$HKU)

    $officeApps = @("Word","Excel","PowerPoint","Outlook","MS Project")

    $HKUManifestKeys = @("SOFTWARE\Wow6432Node\Microsoft\Office",
                         "SOFTWARE\Microsoft\Office",
                         "SOFTWARE\Wow6432Node\Microsoft",
                         "SOFTWARE\Microsoft")

    $HKLMManifestKeys = @("SOFTWARE\Wow6432Node\Microsoft\Office",
                          "SOFTWARE\Microsoft\Office",
                          "SOFTWARE\Wow6432Node\Microsoft",
                          "SOFTWARE\Microsoft"
                          "SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office",
                          "SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft",
                          "SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\USERS\.DEFAULT\Software\Microsoft\Office",
                          "SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\USERS\.DEFAULT\Software\Microsoft")

    
    foreach($hkey in $hkeys){
        switch($hkey){
            '2147483650'{
                foreach($key in $HKLMManifestKeys){
                    if($key -match "Office"){
                        $searchapps = $officeApps
                    } else {
                        $searchApps = "Visio"
                    }
                
                    foreach($app in $searchApps){
                        $path = Join-Path $key $app
                        $fullpath = Join-Path $path "Addins"

                        $enumKeys = $regProv.EnumKey($HKLM, $fullpath)
                        foreach($enumkey in $enumKeys.sNames){
                            if($enumkey -eq $AddinID){
                                $addinpath = Join-Path $fullpath $enumkey
                                $values = $regProv.EnumValues($hklm, $addinpath)
                                foreach($value in $values.sNames){
                                    if($value -eq "Manifest"){
                                        $ManifestValue = ($regProv.GetStringValue($hklm, $addinpath, $value)).sValue
                                        if($ManifestValue -match "|"){
                                            $ManifestValue = $ManifestValue.Split("|")[0]
                                        }

                                        return $ManifestValue;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            '2147483651'{
                $HKUsNames = $regProv.EnumKey($HKU, "")

                foreach($HKUsName in $HKUsNames.sNames){
                    if($HKUsName -notmatch "Default"){
                        foreach($HKUManifestKey in $HKUManifestKeys){
                            $path = Join-Path $HKUsName $HKUManifestKey
                            if($path -match "Office"){
                                $searchapps = $officeApps
                            } else {
                                $searchapps = 'Visio'
                            }

                            foreach($app in $searchapps){
                                $appPath = Join-Path $path $app
                                $addinPath = Join-Path $appPath "Addins"
                                
                                $enumKeys = $regProv.EnumKey($hkey, $addinPath)
                                foreach($enumkey in $enumKeys.sNames){
                                    if($enumkey -eq $AddinID){
                                        $fullpath = Join-Path $addinPath $enumkey
                                        $values = $regProv.EnumValues($hkey, $fullpath)
             
                                        foreach($value in $values.sNames){
                                            if($value -eq "Manifest"){
                                                $ManifestValue = ($regProv.GetStringValue($hkey, $addinpath, $value)).sValue
                                                if($ManifestValue -match "|"){
                                                    $ManifestValue = $ManifestValue.Split("|")[0]
                                                }
                                                 
                                                return $ManifestValue;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}

function Get-AddinLoadtime {
Param(
    [string]$ComputerName = $env:COMPUTERNAME,
    [string]$AddinID
)
    $regProv = Get-WmiObject -List "StdRegProv" -Namespace root\default -ComputerName $ComputerName

    $HKU = [UInt32] "0x80000003"

    $loadTimeKey = "SOFTWARE\Microsoft\Office"
    $officeVersions = @("11.0","12.0","13.0","14.0","15.0","16.0")
    $officeApps = @("Word","Excel","PowerPoint","Outlook","Visio","MS Project")

    $HKUsNames = $regProv.EnumKey($HKU, "")
    
    foreach($HKUsName in $HKUsNames.sNames){
        if($HKUsName -notmatch "Default"){
            $path = Join-Path $HKUsName $loadTimeKey
            foreach($officeVersion in $officeVersions){
                $versionPath = Join-Path $path $officeVersion
                foreach($officeApp in $officeApps){
                    $appPath = Join-Path $versionPath "$officeApp\AddInLoadTimes"

                    $values = $regProv.EnumValues($HKU, $appPath)
                    if($values.sNames.Count -ge 1){
                        foreach($value in $values.sNames){
                            if($value -eq $AddinID){
                                $totalValue = @()
                                $AddinLoadTime = $regProv.GetBinaryValue($HKU, $appPath, $value)

                                foreach($time in $AddinLoadTime.uValue){
                                    $decValue = [convert]::ToString($time, 16)
                                    $decValueCharacters = $decValue | measure -Character
                                    
                                    if($decValueCharacters.Characters -le 1){
                                        $decValue = AddDoubleInt -int $decValue
                                    }
                                
                                    $totalValue += $decValue
                                }
                                
                                $totalValue = [system.string]::Join(" ",$totalValue)
                                
                                return $totalValue;
                                        
                            }
                        }
                    }
                }
            }
        }
    }
}

function AddDoubleInt ($int) {
    $num = "0"
    $num += $int

    return $num;

}

function Get-AddinOfficeVersion {
Param(
    [string]$ComputerName = $env:COMPUTERNAME,
    [string]$AddinID
)
    $regProv = Get-WmiObject -List "StdRegProv" -Namespace root\default -ComputerName $ComputerName

    $HKU = [UInt32] "0x80000003"

    $loadTimeKey = "SOFTWARE\Microsoft\Office"
    $officeVersions = @("11.0","12.0","13.0","14.0","15.0","16.0")
    $officeApps = @("Word","Excel","PowerPoint","Outlook","Visio","MS Project")

    $HKUsNames = $regProv.EnumKey($HKU, "")

    foreach($HKUsName in $HKUsNames.sNames){
        if($HKUsName -notmatch "Default"){
            $path = Join-Path $HKUsName $loadTimeKey
            foreach($officeVersion in $officeVersions){
                $OfficeVersionPath = Join-Path $path $officeVersion
                foreach($officeApp in $officeApps){
                    $officeAppPath = Join-Path $OfficeVersionPath $officeApp
                    $loadTimePath = Join-Path $officeAppPath "AddInLoadTimes"
                    
                    $values = $regProv.EnumValues($HKU, $loadTimePath)
                    foreach($value in $values.sNames){
                        if($value -eq $AddinID){
                            $loadBehaviorValue = $regProv.GetBinaryValue($HKU, $loadTimePath, $value)
                            if($loadBehaviorValue -ne $null){
                                $AddinOfficeVersion = $officeVersion

                                return $AddinOfficeVersion;
                            }
                        }
                    }
                }
            }
        }
    }
}

function Get-OutlookCrashingAddin {
Param(
    [string]$ComputerName = $env:COMPUTERNAME,
    [string]$AddinID
)
    $regProv = Get-WmiObject -List "StdRegProv" -Namespace root\default -ComputerName $ComputerName

    $HKU = [UInt32] "0x80000003"

    $OutlookRegKey = "SOFTWARE\Microsoft\Office"
    $crashingAddinListKey = "Outlook\Resiliency\CrashingAddinList"
    $officeVersions = @("11.0","12.0","13.0","14.0","15.0","16.0")

    $HKUsNames = $regProv.EnumKey($HKU, "")
    
    foreach($HKUsName in $HKUsNames.sNames){
        if($HKUsName -notmatch "Default"){
            $path = Join-Path $HKUsName $OutlookRegKey
            foreach($officeVersion in $officeVersions){
                $offieVersionPath = Join-Path $OutlookRegKey $officeVersion
                $crashingAddinListPath = Join-Path $offieVersionPath $crashingAddinListKey
                $crashingAddinValues =  $regProv.EnumValues($HKU, $crashingAddinListPath)

                foreach($crashingAddinValue in $crashingAddinValues.sNames){
                    if($crashingAddinValue -eq $AddinID){
                        $value = $regProv.GetDWORDValue($HKU, $crashingAddinListPath, $crashingAddinValue)

                        return $value;
                    }
                }
            }
        }
    }
}