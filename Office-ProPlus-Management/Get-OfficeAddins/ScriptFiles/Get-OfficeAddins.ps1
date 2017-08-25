[CmdletBinding(SupportsShouldProcess=$true)]
param(
[Parameter()]
[string]$WmiClassName = "Custom_OfficeAddins"
)

function Get-OfficeAddins {
Param(
    [Parameter()]
    [string]$ComputerName = $env:COMPUTERNAME,

    [Parameter()]
    [string]$WMIClassName = "Custom_OfficeAddins"
)    
    $HKLM = [UInt32] "0x80000002"
    $HKU = [UInt32] "0x80000003"
    
    $officeApps = @("Word","Excel","PowerPoint","Outlook","MS Project")
    
    $HKLMKeys = @("SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office",
                 "SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Visio\Addins",
                 "SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\USERS\.DEFAULT\Software\Microsoft\Office",
                 "SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\USERS\.DEFAULT\Software\Microsoft\Visio\Addins",
                 "Software\Microsoft\Office",  
                 "Software\Wow6432Node\Microsoft\Office",
                 "Software\Microsoft",
                 "Software\Wow6432Node")
    
    $HKUKeys = @("Software\Microsoft\Office",  
                 "Software\Wow6432Node\Microsoft\Office",
                 "Software\Microsoft",
                 "Software\Wow6432Node")
    
    $regProv = Get-WmiObject -List "StdRegProv" -Namespace root\default -ComputerName $ComputerName

    $ClassName = $WMIClassName 
    $classExists = Get-WmiObject -Class $ClassName -ErrorAction SilentlyContinue
    $resiliencyList = Get-ResiliencyAddins
    $OutlookCrashingAddins = Get-OutlookCrashingAddins

    if(!$classExists){
        New-OfficeAddinWMIClass -ClassName $WMIClassName

        New-OfficeAddinWMIProperty -ClassName $WMIClassName

        Set-OfficeAddinWMIPropertyQualifier -ClassName $WMIClassName -PropertyName Name -QualifierName Key -QualifierValue $true
    }
    
    foreach($HKLMKey in $HKLMKeys){
        if($HKLMKey -notmatch "Office"){
            $searchApps = "Visio"
        } else {
            $searchApps = $officeApps
        }

        foreach($officeapp in $searchApps){
            $path = Join-Path $HKLMKey $officeapp
            $hkeyEnum = $regProv.EnumKey($HKLM, $path)
    
            if($hkeyEnum.sNames -contains "Addins"){
                $addinsPath = Join-Path $path "Addins"
                $addinEnum = $regProv.EnumKey($HKLM, $addinsPath)
                foreach($addinapp in $addinEnum.sNames){
                    $addinpath = Join-Path $addinsPath $addinapp
                    $LoadBehavior = ($regprov.GetDWORDValue($HKLM, $addinPath, 'LoadBehavior')).uValue
                    $Description = ($regProv.GetStringValue($HKLM, $addinPath, 'Description')).sValue
                    $FriendlyName = ($regProv.GetStringValue($HKLM, $addinPath, 'FriendlyName')).sValue
                    $FullPath = Get-AddinFullPath -AddinID $addinapp
                    $loadTime = Get-AddinLoadtime -AddinID $addinapp
                    $addinOfficeVersion = Get-AddinOfficeVersion -AddinID $addinapp

                    if(!$Description){
                        $Description = " "
                    }
                    
                    if(!$FriendlyName){
                        $FriendlyName = " "
                    }
                    
                    if(!$FullPath){
                        $FullPath = " "
                    }
                    
                    if(!$loadTime){
                        $loadTime = " "
                    }
                    
                    if(!$addinOfficeVersion){
                        $addinOfficeVersion = " "
                    }
                    
                    if(!$LoadBehavior){
                        $LoadBehavior = " "
                    } else {
                        if(($LoadBehavior -as [string]) -ne $null ){
                            [string]$LoadBehavior = $LoadBehavior
                            $LoadBehaviorProperties = Get-LoadBehavior -name $addinapp -value $LoadBehavior
                        }
                    }
                    
                    if(!$addinpath){
                        $addinpath = " "
                    }

                    $isResilient = $false
                    if($resiliencyList.Name -contains $addinapp){
                        $isResilient = $true
                    }

                    $isOutlookCrashingAddin = $false
                    if($OutlookCrashingAddins -contains $addinapp){
                        $isOutlookCrashingAddin = $true
                    }

                    $ID = New-Guid
     
                    $instanceExists = Get-WMIClassInstance -ClassName $WMIClassName -InstanceName $addinapp
                    if(!$instanceExists){
                        $MyNewInstance = New-OfficeAddinWMIClassInstance -ClassName Custom_OfficeAddins
                    
                        $MyNewInstance.ID = $ID
                        $MyNewInstance.Application = $officeapp
                        $MyNewInstance.ComputerName = $env:COMPUTERNAME
                        $MyNewInstance.Description = $Description
                        $MyNewInstance.FriendlyName = $FriendlyName
                        $MyNewInstance.FullPath = $FullPath
                        $MyNewInstance.LoadBehaviorValue = $LoadBehaviorProperties.Value
                        $MyNewInstance.LoadBehaviorStatus = $LoadBehaviorProperties.Status
                        $MyNewInstance.LoadBehavior = $LoadBehaviorProperties.LoadBehavior
                        $MyNewInstance.LoadTime = $LoadTime
                        $MyNewInstance.Name = $addinapp
                        $MyNewInstance.OfficeVersion = $addinOfficeVersion
                        $MyNewInstance.RegistryPath = $addinpath
                        $MyNewInstance.IsResilient = $isResilient
                        $MyNewInstance.IsOutlookCrashingAddin = $isOutlookCrashingAddin
                        
                        New-OfficeAddinWMIClassInstance -ClassName $ClassName -PutInstance $MyNewInstance
                    } else {
                        Set-InstancePropertyValue -ClassName $WMIClassName -InstanceName $instanceExists -Property Application -PropertyValue $officeapp
                        Set-InstancePropertyValue -ClassName $WMIClassName -InstanceName $instanceExists -Property ComputerName -PropertyValue $env:COMPUTERNAME
                        Set-InstancePropertyValue -ClassName $WMIClassName -InstanceName $instanceExists -Property Description -PropertyValue $Description
                        Set-InstancePropertyValue -ClassName $WMIClassName -InstanceName $instanceExists -Property FriendlyName -PropertyValue $FriendlyName
                        Set-InstancePropertyValue -ClassName $WMIClassName -InstanceName $instanceExists -Property FullPath -PropertyValue $FullPath
                        Set-InstancePropertyValue -ClassName $WMIClassName -InstanceName $instanceExists -Property LoadBehaviorValue -PropertyValue $LoadBehaviorProperties.Value
                        Set-InstancePropertyValue -ClassName $WMIClassName -InstanceName $instanceExists -Property LoadBehaviorStatus -PropertyValue $LoadBehaviorProperties.Status
                        Set-InstancePropertyValue -ClassName $WMIClassName -InstanceName $instanceExists -Property LoadBehavior -PropertyValue $LoadBehaviorProperties.LoadBehavior
                        Set-InstancePropertyValue -ClassName $WMIClassName -InstanceName $instanceExists -Property LoadTime -PropertyValue $LoadTime
                        Set-InstancePropertyValue -ClassName $WMIClassName -InstanceName $instanceExists -Property Name -PropertyValue $addinapp
                        Set-InstancePropertyValue -ClassName $WMIClassName -InstanceName $instanceExists -Property OfficeVersion -PropertyValue $addinOfficeVersion
                        Set-InstancePropertyValue -ClassName $WMIClassName -InstanceName $instanceExists -Property RegistryPath -PropertyValue $addinpath
                        Set-InstancePropertyValue -ClassName $WMIClassName -InstanceName $instanceExists -Property IsResilient -PropertyValue $isResilient
                        Set-InstancePropertyValue -ClassName $WMIClassName -InstanceName $instanceExists -Property IsOutlookCrashingAddin -PropertyValue $isOutlookCrashingAddin
                    }
                }
            }
        }
    }

    foreach($HKUKey in $HKUKeys){
        if($HKUKey -notmatch "Office"){
            $searchApps = "Visio"
        } else {
            $searchApps = $officeApps
        }

        $HKUsNames = $regProv.EnumKey($HKU, "")

        foreach($HKUsName in $HKUsNames.sNames){
            if($HKUsName -notmatch "Default"){
                $HKUPath = Join-Path $HKUsName $HKUKey 

                foreach($officeapp in $searchApps){
                    $path = Join-Path $HKUPath $officeapp 
                    $hkeyEnum = $regProv.EnumKey($HKU, $path)
        
                    if($hkeyEnum.sNames -contains "Addins"){
                        $addinsPath = Join-Path $path "Addins"
                        $addinEnum = $regProv.EnumKey($HKU, $addinsPath)
                        foreach($addinapp in $addinEnum.sNames){
                            $addinpath = Join-Path $addinsPath $addinapp
                            $LoadBehavior = ($regprov.GetDWORDValue($HKU, $addinPath, 'LoadBehavior')).uValue
                            $Description = ($regProv.GetStringValue($HKU, $addinPath, 'Description')).sValue
                            $FriendlyName = ($regProv.GetStringValue($HKU, $addinPath, 'FriendlyName')).sValue
                            $FullPath = Get-AddinFullPath -AddinID $addinapp -AddinType "VSTO"
                            $loadTime = Get-AddinLoadtime -AddinID $addinapp
                            $addinOfficeVersion = Get-AddinOfficeVersion -AddinID $addinapp
                    
                            if(!$Description){
                                $Description = " "
                            }
                            
                            if(!$FriendlyName){
                                $FriendlyName = " "
                            }
                            
                            if(!$FullPath){
                                $FullPath = " "
                            }
                            
                            if(!$loadTime){
                                $loadTime = " "
                            }
                            
                            if(!$addinOfficeVersion){
                                $addinOfficeVersion = " "
                            }
                            
                            if(!$LoadBehavior){
                                $LoadBehavior = " "
                            } else {
                                if(($LoadBehavior -as [string]) -ne $null ){
                                    [string]$LoadBehavior = $LoadBehavior
                                    $LoadBehaviorProperties = Get-LoadBehavior -name $addinapp -value $LoadBehavior
                                }
                            }
                            
                            if(!$addinpath){
                                $addinpath = " "
                            }

                            $isResilient = $false
                            if($resiliencyList.Name -contains $addinapp){
                                $isResilient = $true
                            }

                            $isOutlookCrashingAddin = $false
                            if($OutlookCrashingAddins -contains $addinapp){
                                $isOutlookCrashingAddin = $true
                            }

                            $ID = New-Guid

                            $instanceExists = Get-WMIClassInstance -ClassName $WMIClassName -InstanceName $addinapp
                            if(!$instanceExists){
                                $MyNewInstance = New-OfficeAddinWMIClassInstance -ClassName Custom_OfficeAddins
                            
                                $MyNewInstance.ID = $ID
                                $MyNewInstance.Application = $officeapp
                                $MyNewInstance.ComputerName = $env:COMPUTERNAME
                                $MyNewInstance.Description = $Description
                                $MyNewInstance.FriendlyName = $FriendlyName
                                $MyNewInstance.FullPath = $FullPath
                                $MyNewInstance.LoadBehaviorValue = $LoadBehaviorProperties.Value
                                $MyNewInstance.LoadBehaviorStatus = $LoadBehaviorProperties.Status
                                $MyNewInstance.LoadBehavior = $LoadBehaviorProperties.LoadBehavior
                                $MyNewInstance.LoadTime = $LoadTime
                                $MyNewInstance.Name = $addinapp
                                $MyNewInstance.OfficeVersion = $addinOfficeVersion
                                $MyNewInstance.RegistryPath = $addinpath
                                $MyNewInstance.IsResilient = $isResilient
                                $MyNewInstance.IsOutlookCrashingAddin = $isOutlookCrashingAddin
                                
                                New-OfficeAddinWMIClassInstance -ClassName $ClassName -PutInstance $MyNewInstance
                            } else {
                                Set-InstancePropertyValue -ClassName $WMIClassName -InstanceName $instanceExists -Property Application -PropertyValue $officeapp
                                Set-InstancePropertyValue -ClassName $WMIClassName -InstanceName $instanceExists -Property ComputerName -PropertyValue $env:COMPUTERNAME
                                Set-InstancePropertyValue -ClassName $WMIClassName -InstanceName $instanceExists -Property Description -PropertyValue $Description
                                Set-InstancePropertyValue -ClassName $WMIClassName -InstanceName $instanceExists -Property FriendlyName -PropertyValue $FriendlyName
                                Set-InstancePropertyValue -ClassName $WMIClassName -InstanceName $instanceExists -Property FullPath -PropertyValue $FullPath
                                Set-InstancePropertyValue -ClassName $WMIClassName -InstanceName $instanceExists -Property LoadBehaviorValue -PropertyValue $LoadBehaviorProperties.Value
                                Set-InstancePropertyValue -ClassName $WMIClassName -InstanceName $instanceExists -Property LoadBehaviorStatus -PropertyValue $LoadBehaviorProperties.Status
                                Set-InstancePropertyValue -ClassName $WMIClassName -InstanceName $instanceExists -Property LoadBehavior -PropertyValue $LoadBehaviorProperties.LoadBehavior
                                Set-InstancePropertyValue -ClassName $WMIClassName -InstanceName $instanceExists -Property LoadTime -PropertyValue $LoadTime
                                Set-InstancePropertyValue -ClassName $WMIClassName -InstanceName $instanceExists -Property Name -PropertyValue $addinapp
                                Set-InstancePropertyValue -ClassName $WMIClassName -InstanceName $instanceExists -Property OfficeVersion -PropertyValue $addinOfficeVersion
                                Set-InstancePropertyValue -ClassName $WMIClassName -InstanceName $instanceExists -Property RegistryPath -PropertyValue $addinpath
                                Set-InstancePropertyValue -ClassName $WMIClassName -InstanceName $instanceExists -Property IsResilient -PropertyValue $isResilient
                                Set-InstancePropertyValue -ClassName $WMIClassName -InstanceName $instanceExists -Property IsOutlookCrashingAddin -PropertyValue $isOutlookCrashingAddin
                            }
                        }
                    }
                }
            }
        }
    }
}

function Get-AddinFullPath {
Param(
    [string]$ComputerName = $env:COMPUTERNAME,
    [string]$AddinID
)

    $regProv = Get-WmiObject -List "StdRegProv" -Namespace root\default -ComputerName $ComputerName

    $HKLM = [UInt32] "0x80000002"

    $clsidPathKeys = @("SOFTWARE\Classes\CLSID",
                     "SOFTWARE\Classes\Wow6432Node\CLSID",
                     "SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\SOFTWARE\Classes\CLSID",
                     "SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\SOFTWARE\Classes\Wow6432Node\CLSID")

    $manifestKey = Get-ManifestKey -AddinID $AddinID
    if($manifestKey -ne $null){
        return $manifestKey
    } else {
        $clsid = Get-CLSID -ProgId $AddinID
        foreach($key in $clsidPathKeys){
            $path = Join-Path $key $clsid
            $InProcPath = Join-Path $path "InprocServer32"
            if(Test-Path "HKLM:\$InProcPath"){
                $fullpath = Get-ItemProperty ("HKLM:\$InProcPath")
                $fullpath = $fullpath.'(default)'
              
                return $fullpath
            }
        }
    }
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

                                if(($totalValue -as [string]) -ne $null ){
                                    [string]$totalValue = $totalValue
                                }
                                
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

function Get-OutlookCrashingAddins {
Param(
    [string]$ComputerName = $env:COMPUTERNAME
)
   
    $regProv = Get-WmiObject -List "StdRegProv" -Namespace root\default -ComputerName $ComputerName

    $HKU = [UInt32] "0x80000003"

    $OutlookRegKey = "SOFTWARE\Microsoft\Office"
    $crashingAddinListKey = "Outlook\Resiliency\CrashingAddinList"
    $officeVersions = @("11.0","12.0","13.0","14.0","15.0","16.0")

    $HKUsNames = $regProv.EnumKey($HKU, "")
    
    $CrashingAddinList = @()
    foreach($HKUsName in $HKUsNames.sNames){
        if($HKUsName -notmatch "Default"){
            $path = Join-Path $HKUsName $OutlookRegKey
            foreach($officeVersion in $officeVersions){
                $officeVersionPath = Join-Path $OutlookRegKey $officeVersion
                $crashingAddinListPath = Join-Path $officeVersionPath $crashingAddinListKey
                $crashingAddinValues =  $regProv.EnumValues($HKU, $crashingAddinListPath)
                foreach($crashingAddinValue in $crashingAddinValues.sNames){
                    $value = $regProv.GetDWORDValue($HKU, $crashingAddinListPath, $crashingAddinValue)
                    if($value.uValue -eq "1"){
                        if($CrashingAddinList -notcontains $crashingAddinValue){
                            $CrashingAddinList += $crashingAddinValue
                        }
                    }
                }
            }
        }
    }

    return $CrashingAddinList;
}

function Get-ResiliencyAddins{
Param(
    [string]$ComputerName = $env:COMPUTERNAME
)
    $defaultDisplaySet = 'Name','Value', 'Status'
    $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet',[string[]]$defaultDisplaySet)
    $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
    $results = New-Object PSObject[] 0;
    $regProv = Get-WmiObject -List "StdRegProv" -Namespace root\default -ComputerName $ComputerName

    $HKU = [UInt32] "0x80000003"

    $keyStart = "Software\Microsoft\Office"
    $officeApps = @("Word","Excel","PowerPoint","Outlook","MS Project")
    $keyEnd = "Resiliency\DoNotDisableAddinList"
    $officeVersions = @("11.0","12.0","13.0","14.0","15.0","16.0")

    $HKUsNames = $regProv.EnumKey($HKU, "")
    
    $resiliencyList = @()
    foreach($HKUsName in $HKUsNames.sNames){
        if($HKUsName -notmatch "Default"){
            foreach($officeVersion in $officeVersions){
                $StartPath = Join-Path $HKUsName $keyStart
                $officeVersionPath = Join-Path $StartPath $officeVersion
                foreach($officeApp in $officeApps){
                    $appPath = Join-Path $officeVersionPath $officeApp
                    $fullpath = Join-Path $appPath $keyEnd                    
                    $values = $regProv.EnumValues($HKU, $fullpath)
                    if($values.sNames){
                        foreach($value in $values.sNames){
                            if($resiliencyList -notcontains $value){
                                $resiliencyList += $value
                                $dwordValue = $regProv.GetDWORDValue($HKU, $fullpath, $value)

                                switch($dwordValue.uValue){
                                    "1"{
                                        $Status = "Boot load"
                                    }
                                    "2"{
                                        $Status = "Demand load"
                                    }
                                    "3"{
                                        $Status = "Crash"
                                    }
                                    "4"{
                                        $Status = "Handling FolderSwitch event"
                                    }
                                    "5"{
                                        $Status = "Handling BeforeFolderSwitch event"
                                    }
                                    "6"{
                                        $Status = "Item Open"
                                    }
                                    "7"{
                                        $Status = "Iteration Count"
                                    }
                                    "8"{
                                        $Status = "Shutdown"
                                    }
                                    "9"{
                                        $Status = "Crash, but not disabled because add-in is in the allow list"
                                    }
                                    "10"{
                                        $Status = "Crash, but not disabled because user selected no in disable dialog"
                                    }
                                }
                           
                                $object = New-Object PSObject -Property @{Name = $value; Value = $dwordValue.uValue; Status = $Status}
                                $object | Add-Member MemberSet PSStandardMembers $PSStandardMembers
                                $results += $object
                            }
                        }
                    }
                }
            }
        }
    }

    return $results;
}

function New-OfficeAddinWMIClass{
Param(
    [Parameter()]
    [string]$ClassName = "Custom_OfficeAddins",

    [Parameter()]
    [string]$NameSpace = "root\cimv2"
)
    $NewClass = New-Object System.Management.ManagementClass($NameSpace, $null, $null)
    $NewClass.Name = $ClassName
    $NewClass.Put() | Out-Null
}

Function New-OfficeAddinWMIProperty{
[CmdletBinding()]
	Param(
		[Parameter()]
        [string]$ClassName = "Custom_OfficeAddins",

        [Parameter()]
        [string]$NameSpace="Root\cimv2",

        [Parameter()]
        [string[]]$PropertyName,

        [Parameter()]
        [string]$PropertyValue = ""
	)
    
    [wmiclass]$OfficeAddinWMIClass = Get-WmiObject -Class $ClassName -Namespace $NameSpace -list
    if(!$PropertyName){
        $PropertyName = @("ID", "Application", "ComputerName", "Description", "FriendlyName", "FullPath", "LoadBehaviorValue", 
                          "LoadBehaviorStatus", "LoadBehavior", "LoadTime","Name", "OfficeVersion", "RegistryPath", "IsResilient", "IsOutlookCrashingAddin")
    }
   
    foreach($property in $PropertyName){
        $OfficeAddinWMIClass.Properties.add($property,$PropertyValue)
        $OfficeAddinWMIClass.Put() | Out-Null
    }                                          
}

Function Set-OfficeAddinWMIPropertyQualifier{
[CmdletBinding()]
Param(
	[Parameter()]
       [string]$ClassName = "Custom_OfficeAddins",

       [Parameter()]
       [string]$NameSpace="Root\cimv2",

       [Parameter()]
       [string]$PropertyName = "Name",

       [Parameter()]
       $QualifierName = "Key",

       [Parameter()]
       $QualifierValue = $true,

       [switch]$key,
       [switch]$IsAmended = $false,
       [switch]$IsLocal = $true,
       [switch]$PropagatesToInstance = $true,
       [switch]$PropagesToSubClass = $false,
       [switch]$IsOverridable = $true
)

    $OfficeAddinWmiClass = Get-OfficeAddinWMIClass -ClassName $ClassName -NameSpace $NameSpace

    if($OfficeAddinWmiClass.Properties[$PropertyName]){    
        if ($Key){
            $OfficeAddinWmiClass.Properties[$PropertyName].Qualifiers.Add("Key",$true)
            $OfficeAddinWmiClass.put() | out-null
        }else{
            $OfficeAddinWmiClass.Properties[$PropertyName].Qualifiers.add($QualifierName,$QualifierValue, $IsAmended,$IsLocal,$PropagatesToInstance,$PropagesToSubClass)
            $OfficeAddinWmiClass.put() | out-null
        }
    }
}

Function Get-OfficeAddinWMIProperty{
[CmdletBinding()]
Param(
	[Parameter()]
       [string]$ClassName = "Custom_OfficeAddins",

       [Parameter()]
       [string]$NameSpace="Root\cimv2",

       [Parameter()]
       [string]$PropertyName
)
    if($PropertyName){
        $return = (Get-OfficeAddinWMIClass -ClassName $ClassName -NameSpace $NameSpace ).properties["$($PropertyName)"]
    }else{
        $return = (Get-OfficeAddinWMIClass -ClassName $ClassName -NameSpace $NameSpace ).properties            
    } 
      
    Return $return      
}

Function Get-OfficeAddinWMIClass{
[CmdletBinding()]
	Param(
		[Parameter(ValueFromPipeLine=$true)]
        [string]$ClassName,

        [Parameter()]
        [string]$NameSpace = "root\cimv2"
	)  
    
    if (!($ClassName)){
        $return = Get-WmiObject -Namespace $NameSpace -Class * -list
    }else{
        $return = Get-WmiObject -Namespace $NameSpace -Class $ClassName -list
    }
    
    return $return
}

Function New-OfficeAddinWMIClassInstance {
[CmdletBinding()]
Param(
	[Parameter(Mandatory=$true)]
    [string]$ClassName,

    [Parameter(Mandatory=$false)]
    [string]$NameSpace="Root\cimv2",

    [Parameter(valueFromPipeLine=$true)]$PutInstance
)
    
    $WmiClass = Get-OfficeAddinWMIClass -NameSpace $NameSpace -ClassName $ClassName
     
    if($PutInstance){  
        $PutInstance.Put()
    }else{
        $CreateInstance = $WmiClass.CreateInstance()
        $CreateInstance
    }       
}

function Set-InstancePropertyValue{
Param(
    [string]$ClassName,

    [string]$InstanceName,

    [string]$Property,

    [string]$PropertyValue
)
    
    [wmiclass]$WmiClass = Get-WmiObject -Class $ClassName -List

    $instance = $WmiClass.GetInstances() | ? {$_.Name -eq $InstanceName}

    $instance.SetPropertyValue($Property, $PropertyValue)

}

function Get-WMIClassInstance{
Param(
    [string]$ClassName,

    [string]$InstanceName
)
    
    [wmiclass]$WmiClass = Get-WmiObject -Class $ClassName -List

    $instance = $WmiClass.GetInstances() | ? {$_.Name -eq $InstanceName}

    return $instance.Name

}

function Get-LoadBehavior{
Param(
    [string]$name,
    [string]$value
)
    $defaultDisplaySet = 'Name','Value', 'Status', 'LoadBehavior'

    $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet',[string[]]$defaultDisplaySet)
    $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)

    $results = new-object PSObject[] 0;

    switch($value){
        "0"{
            $status = "Unloaded"
            $LoadBehavior = "Do not load automatically"
        }
        "1"{
            $status = "Loaded"
            $LoadBehavior = "Do not load automatically"
        }
        "2"{
            $status = "Unloaded"
            $LoadBehavior = "Load at startup"
        }
        "3"{
            $status = "Loaded"
            $LoadBehavior = "Load at startup"
        }
        "8"{
            $status = "Unloaded"
            $LoadBehavior = "Load on demand"
        }
        "9"{
            $status = "Loaded"
            $LoadBehavior = "Load on demand"
        }
        "16"{
            $status = "Loaded"
            $LoadBehavior = "Load first time, then load on demand"
        }
    }

    $object = New-Object PSObject -Property @{Name = $name; Value = $value; Status = $status; LoadBehavior = $LoadBehavior}
    $object | Add-Member MemberSet PSStandardMembers $PSStandardMembers
    $results += $object

    return $results
}

function New-GUID{
 $guid = [guid]::NewGuid()

 return $guid
}

Function IsDotSourced() {
  [CmdletBinding(SupportsShouldProcess=$true)]
  param(
    [Parameter(ValueFromPipelineByPropertyName=$true)]
    [string]$InvocationLine = ""
  )
  $cmdLine = $InvocationLine.Trim()
  Do {
    $cmdLine = $cmdLine.Replace(" ", "")
  } while($cmdLine.Contains(" "))

  $dotSourced = $false
  if ($cmdLine -match '^\.\\') {
     $dotSourced = $false
  } else {
     $dotSourced = ($cmdLine -match '^\.')
  }

  return $dotSourced
}

$dotSourced = IsDotSourced -InvocationLine $MyInvocation.Line

if (!($dotSourced)) {
    Get-OfficeAddins -WMIClassName $WmiClassName
}