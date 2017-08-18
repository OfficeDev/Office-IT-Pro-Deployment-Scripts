function Get-OfficeAddins {
Param(
    [string]$ComputerName = $env:COMPUTERNAME
)

    $defaultDisplaySet = 'ComputerName','Application','Name','Description','FriendlyName','LoadBehavior','RegistryPath','FullPath','LoadTime','OfficeVersion'
    $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet',[string[]]$defaultDisplaySet)
    $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
    $results = New-Object PSObject[] 0;
    
    $HKCU = [UInt32] "0x80000001"
    $HKLM = [UInt32] "0x80000002"
    $HKU = [UInt32] "0x80000003"
    
    $HKEYS = @($HKCU, $HKLM)
    
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

    $ClassName = "Custom_OfficeAddins" 
    $classExists = Get-WmiObject -Class $ClassName -ErrorAction SilentlyContinue
    if(!$classExists){
        New-OfficeAddinWMIClass -ClassName $ClassName

        New-OfficeAddinWMIProperty -ClassName $ClassName

        Set-OfficeAddinWMIPropertyQualifier -ClassName $ClassName -PropertyName Name -QualifierName Key -QualifierValue $true
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
                        }
                    }
                    
                    if(!$addinpath){
                        $addinpath = " "
                    }

                    $MyNewInstance = New-OfficeAddinWMIClassInstance -ClassName Custom_OfficeAddins
                    
                    $MyNewInstance.Application = $officeapp
                    $MyNewInstance.ComputerName = $env:COMPUTERNAME
                    $MyNewInstance.Description = $Description
                    $MyNewInstance.FriendlyName = $FriendlyName
                    $MyNewInstance.FullPath = $FullPath
                    $MyNewInstance.LoadBehavior = $LoadBehavior
                    $MyNewInstance.LoadTime = $LoadTime
                    $MyNewInstance.Name = $addinapp
                    $MyNewInstance.OfficeVersion = $addinOfficeVersion
                    $MyNewInstance.RegistryPath = $addinpath
                    
                    New-OfficeAddinWMIClassInstance -ClassName $ClassName -PutInstance $MyNewInstance
                    
                    
                    #New-CimInstance -ClassName Custom_OfficeAddins -Property @{Application=$officeapp; ComputerName=$ComputerName; Description=$Description; FriendlyName=$FriendlyName; FullPath=$FullPath;
                    #                                                  LoadBehavior=$LoadBehavior; LoadTime=$loadTime; Name=$addinapp; OfficeVersion=$addinOfficeVersion; RegistryPath=$addinpath}
    
                    #$object = New-Object PSObject -Property @{ComputerName = $ComputerName; Application = $officeapp; Name = $addinapp; RegistryPath = $addinpath; 
                    #                                          Description = $Description; FriendlyName = $FriendlyName; LoadBehavior = $LoadBehavior;
                    #                                          FullPath = $FullPath; LoadTime = $loadTime; OfficeVersion = $addinOfficeVersion}
                    #$object | Add-Member MemberSet PSStandardMembers $PSStandardMembers
                    #$results += $object
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
                                }
                            }
                            
                            if(!$addinpath){
                                $addinpath = " "
                            }

                            $MyNewInstance = New-OfficeAddinWMIClassInstance -ClassName Custom_OfficeAddins
                    
                            $MyNewInstance.Application = $officeapp
                            $MyNewInstance.ComputerName = $env:COMPUTERNAME
                            $MyNewInstance.Description = $Description
                            $MyNewInstance.FriendlyName = $FriendlyName
                            $MyNewInstance.FullPath = $FullPath
                            $MyNewInstance.LoadBehavior = $LoadBehavior
                            $MyNewInstance.LoadTime = $LoadTime
                            $MyNewInstance.Name = $addinapp
                            $MyNewInstance.OfficeVersion = $addinOfficeVersion
                            $MyNewInstance.RegistryPath = $addinpath
                            
                            New-OfficeAddinWMIClassInstance -ClassName $ClassName -PutInstance $MyNewInstance
                            
                            #New-CimInstance -ClassName Custom_OfficeAddins -Property @{Application=$officeapp; ComputerName=$ComputerName; Description=$Description; FriendlyName=$FriendlyName; FullPath=$FullPath;
                            #                                                           LoadBehavior=$LoadBehavior; LoadTime=$loadTime; Name=$addinapp; OfficeVersion=$addinOfficeVersion; RegistryPath=$addinpath}
        
                            #$object = New-Object PSObject -Property @{ComputerName = $ComputerName; Application = $officeapp; Name = $addinapp; RegistryPath = $addinpath;
                            #                                          Description = $Description; FriendlyName = $FriendlyName; LoadBehavior = $LoadBehavior;
                            #                                          FullPath = $FullPath; LoadTime = $loadTime; OfficeVersion = $addinOfficeVersion}
                            #$object | Add-Member MemberSet PSStandardMembers $PSStandardMembers
                            #$results += $object
                        }
                    }
                }
            }
        }
    }
    
    #return $results;

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
                $officeVersionPath = Join-Path $OutlookRegKey $officeVersion
                $crashingAddinListPath = Join-Path $officeVersionPath $crashingAddinListKey
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

function New-CustomOfficeAddinWMIClass{
    $newClass = New-Object System.Management.ManagementClass ("root\cimv2", [String]::Empty, $null); 
    
    $newClass["__CLASS"] = "Custom_OfficeAddins"; 
    
    $newClass.Qualifiers.Add("Static", $true)
    $newClass.Properties.Add("ComputerName", [System.Management.CimType]::String, $false)
    $newClass.Properties["ComputerName"].Qualifiers.Add("Key", $true)
    $newClass.Properties.Add("Application", [System.Management.CimType]::String, $false)
    $newClass.Properties["Application"].Qualifiers.Add("Key", $true)
    $newClass.Properties.Add("Name", [System.Management.CimType]::String, $false)
    $newClass.Properties["Name"].Qualifiers.Add("Key", $true)
    $newClass.Properties.Add("Description", [System.Management.CimType]::String, $false)
    $newClass.Properties["Description"].Qualifiers.Add("Key", $true)
    $newClass.Properties.Add("FriendlyName", [System.Management.CimType]::String, $false)
    $newClass.Properties["FriendlyName"].Qualifiers.Add("Key", $true)
    $newClass.Properties.Add("LoadBehavior", [System.Management.CimType]::String, $false)
    $newClass.Properties["LoadBehavior"].Qualifiers.Add("Key", $true)
    $newClass.Properties.Add("RegistryPath", [System.Management.CimType]::String, $false)
    $newClass.Properties["RegistryPath"].Qualifiers.Add("Key", $true)
    $newClass.Properties.Add("FullPath", [System.Management.CimType]::String, $false)
    $newClass.Properties["FullPath"].Qualifiers.Add("Key", $true)
    $newClass.Properties.Add("LoadTime", [System.Management.CimType]::String, $false)
    $newClass.Properties["LoadTime"].Qualifiers.Add("Key", $true)
    $newClass.Properties.Add("OfficeVersion", [System.Management.CimType]::String, $false)
    $newClass.Properties["OfficeVersion"].Qualifiers.Add("Key", $true)
    
    $newClass.Put()
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
        $PropertyName = @("Application", "ComputerName", "Description", "FriendlyName", "FullPath", "LoadBehavior", "LoadTime","Name", "OfficeVersion", "RegistryPath")
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