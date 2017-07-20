function Get-OfficeAddins {
Param(
    [string]$ComputerName = $env:COMPUTERNAME
)

$defaultDisplaySet = 'Application','Name','Description','FriendlyName','LoadBehavior'
$defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet',[string[]]$defaultDisplaySet)
$PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
$results = New-Object PSObject[] 0;

$HKCU = [UInt32] "0x80000001"
$HKLM = [UInt32] "0x80000002"

$HKEYS = @($HKCU, $HKLM)

$officeApps = @("Word","Excel","PowerPoint","Outlook","MS Project")

$officeKeys = @("SOFTWARE\Microsoft\Office",
                "SOFTWARE\Wow6432Node\Microsoft\Office",
                "SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office",
                "SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\USER\.DEFAULT\Software\Microsoft\Office",
                "SOFTWARE\Microsoft\Visio",
                "SOFTWARE\Wow6432Node\Microsoft\Visio",
                "SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Visio",
                "SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\USER\.DEFAULT\Software\Microsoft\Visio")

$COMKeys = @("SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office",
             "SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Visio\Addins",
             "SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\USER\.DEFAULT\Software\Microsoft\Office",
             "SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\USER\.DEFAULT\Software\Microsoft\Visio\Addins")

$VSTOKeys = @("Software\Microsoft\Office\application name\Addins",  
              "Software\Wow6432Node\Microsoft\Office\application name\Addins",
              "Software\Microsoft\Visio\Addins",
              "Software\Wow6432Node\Visio\Addins")

$regProv = Get-WmiObject -List "StdRegProv" -Namespace root\default -ComputerName $ComputerName

foreach($officekey in $officeKeys){
    foreach($officeapp in $officeApps){
        $path = Join-Path $officekey $officeapp

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
                  
                    $LoadBehavior = $regprov.GetDWORDValue($hkey, $addinPath, 'LoadBehavior')
                    $Description = $regProv.GetStringValue($hkey, $addinPath, 'Description')
                    $FriendlyName = $regProv.GetStringValue($hkey, $addinPath, 'FriendlyName')

                    $object = New-Object PSObject -Property @{Application = $officeapp; Hive = $hive; Name = $addinapp; Path = $addinpath; Description = $Description.sValue; FriendlyName = $FriendlyName.sValue; LoadBehavior = $LoadBehavior.uValue}
                    $object | Add-Member MemberSet PSStandardMembers $PSStandardMembers
                    $results += $object
                }
            }
        }
    }
}

return $results;

}

function Get-RegistryKeys {
Param(
    [string]$ComputerName = $env:COMPUTERNAME,
    [string]$KeyPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\Outlook\Addins"
) 
    $defaultDisplaySet = 'Path','ValueName','ValueEntry'
    $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet',[string[]]$defaultDisplaySet)
    $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
    $results = New-Object PSObject[] 0;

    $regProv = Get-Wmiobject -List "StdRegProv" -Namespace root\default -ComputerName $ComputerName
    $Hive = $KeyPath.Split("\")[0]

    switch($Hive) {
        "HKEY_CLASSES_ROOT" {
            $HKEY = [UInt32] "0x80000000"
            $ShortKeyPath = $KeyPath.Trim("HKEY_CLASSES_ROOT\")
        }
        "HKEY_CURRENT_USER" {
            $HKEY = [UInt32] "0x80000001"
            $ShortKeyPath = $KeyPath.Trim("HKEY_CURRENT_USER\")
        }
        "HKEY_LOCAL_MACHINE" {
            $HKEY = [UInt32] "0x80000002"
            $ShortKeyPath = $KeyPath.Trim("HKEY_LOCAL_MACHINE\")
        }
        "HKEY_USERS" {
            $HKEY = [UInt32] "0x80000003"
            $ShortKeyPath = $KeyPath.Trim("HKEY_USERS\")
        }
        "HKEY_CURRENT_CONFIG" {
            $HKEY = [UInt32] "0x80000004"
            $ShortKeyPath = $KeyPath.Trim("HKEY_CURRENT_CONFIG\")
        }
    }

    # Get a list of the subkeys
    $subKeys = ($regProv.EnumKey($HKEY,$ShortKeyPath)).sNames
    if($subKeys.GetType().Name -eq "String[]"){
        foreach($key in $subKeys){
            $Path = Join-Path $ShortKeyPath $key
            $Values = $regProv.EnumValues($HKEY, $Path)
            if($Values.sNames.Count -gt '0'){
                foreach($val in $Values.sNames){
                    $fullPath = Join-Path $KeyPath $key
                    $valueType = Get-RegistryValueType -Path $fullPath -Value $val

                    switch($valueType){
                        "String"{
                            $DataValue = $regProv.GetStringValue($HKEY, $Path, $val)
                            $Value = $DataValue.sValue
                        }
                        "DWORD"{
                            $DataValue = $regProv.GetDWORDValue($HKEY, $Path, $val)
                            $Value = $DataValue.uValue
                        }
                        "QWORD"{
                            $DataValue = $regProv.GetQWORDValue($HKEY, $Path, $val)
                            $Value = $DataValue.uValue
                        }
                        "BINARY"{
                            $DataValue = $regProv.GetBinaryValue($HKEY, $Path, $val)
                            $Value = $DataValue.uValue
                        }
                        "MULTI-STRING"{
                            $DataValue = $regProv.GetMultiStringValue($HKEY, $Path, $val)
                            $Value = $DataValue.sValue
                        }
                    }

                    $object = New-Object PSObject -Property @{Path = $Path; ValueName = $val; ValueEntry = $Value}
                    $object | Add-Member MemberSet PSStandardMembers $PSStandardMembers
                    $results += $object
                       
                }
            } else {

            }
        }
    } else {

    }

    return $results
}

function Get-OfficeAddIns2 {
Param(
    [string]$ComputerName = $env:COMPUTERNAME,
    [string]$KeyPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\Outlook\Addins"
) 
    $defaultDisplaySet = 'Name','Path'
    $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet',[string[]]$defaultDisplaySet)
    $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
    $results = New-Object PSObject[] 0;
        
    $regProv = Get-WmiObject -List "StdRegProv" -Namespace root\default -ComputerName $ComputerName
    $Hive = $KeyPath.Split("\")[0]

    switch($Hive) {
        "HKEY_CLASSES_ROOT" {
            $HKEY = [UInt32] "0x80000000"
            $ShortKeyPath = $KeyPath.Trim("HKEY_CLASSES_ROOT\")
        }
        "HKEY_CURRENT_USER" {
            $HKEY = [UInt32] "0x80000001"
            $ShortKeyPath = $KeyPath.Trim("HKEY_CURRENT_USER\")
        }
        "HKEY_LOCAL_MACHINE" {
            $HKEY = [UInt32] "0x80000002"
            $ShortKeyPath = $KeyPath.Trim("HKEY_LOCAL_MACHINE\")
        }
        "HKEY_USERS" {
            $HKEY = [UInt32] "0x80000003"
            $ShortKeyPath = $KeyPath.Trim("HKEY_USERS\")
        }
        "HKEY_CURRENT_CONFIG" {
            $HKEY = [UInt32] "0x80000004"
            $ShortKeyPath = $KeyPath.Trim("HKEY_CURRENT_CONFIG\")
        }
    }

    # Get a list of the addins under the KeyPath
    $subKeys = ($regProv.EnumKey($HKEY,$ShortKeyPath)).sNames
    foreach($key in $subKeys){
        $fullPath = Join-Path $KeyPath $key
        $ShortPathName = Join-Path $ShortKeyPath $key
        $values = $regProv.EnumValues($HKEY, $ShortPathName)


        $object = New-Object PSObject -Property @{Name = $key; Path = $KeyPath; FullPath = $fullPath; Values = $Values.sNames}
        $object | Add-Member MemberSet PSStandardMembers $PSStandardMembers
        $results += $object
    }

    return $results
}

function Get-RegistryValueType {
Param(
    [string]$Path,
    [string]$Value
)  
    $hive = Get-RegistryHive -Path $Path
    $shortPath = Get-ShortKeyPath $Path
    $fullPath = $hive + ":\" + $shortPath
    $valueType = (Get-ItemProperty $fullPath).$Value.GetType()

    switch($valueType.Name){
        "String"{
            $ValueType = "String"
        }
        "Int32"{
            $ValueType = "DWORD"
        }
        "Int64"{
            $ValueType = "QWORD"
        }
        "Byte[]"{
            $ValueType = "BINARY"
        }
        "String[]"{
            $ValueType = "MULTI-STRING"
        }
    }

    return $valueType
}

function Get-RegistryHive($Path) {

    $Hive = $Path.Split("\")[0]

    switch($Hive) {
        "HKEY_CLASSES_ROOT" {
            $HKEY = "HKCR"
        }
        "HKEY_CURRENT_USER" {
            $HKEY = "HKCU"
        }
        "HKEY_LOCAL_MACHINE" {
            $HKEY = "HKLM"
        }
        "HKEY_USERS" {
            $HKEY = "HKU"
        }
        "HKEY_CURRENT_CONFIG" {
            $HKEY = "HKCC"
        }
        "HKCR:" {
            $HKEY = "HKCR"
        }
        "HKCU:" {
            $HKEY = "HKCU"
        }
        "HKLM:" {
            $HKEY = "HKLM"
        }
        "HKCC" {
            $HKEY = "HKU"
        }
    }

    return $HKEY
}

function Get-ShortKeyPath($Path) {
    
    $Hive = $Path.Split("\")[0]

    $shortPath = $Path.Trim("$Hive\")

    return $shortPath
}

function Get-OfficeAddin {
Param(
    [string]$Name
)

$addin = Get-OfficeAddIns | ? {$_.Name -match "Name"}

}

$officeApps = @("Word","Excel","PowerPoint","Outlook","MS Project")

$officeKeys = @("SOFTWARE\Microsoft\Office",
                "SOFTWARE\Wow6432Node\Microsoft\Office",
                "SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office",
                "SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\USER\.DEFAULT\Software\Microsoft\Office")
$visioKeys = @("SOFTWARE\Microsoft\Visio",
               "SOFTWARE\Wow6432Node\Microsoft\Visio",
               "Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Visio",
               "Microsoft\Office\ClickToRun\REGISTRY\USER\.DEFAULT\Software\Microsoft\Visio")



<#
#AddInId 
HKCU\SOFTWARE\Microsoft\Office\[Word|Excel|PowerPoint|Outlook|MS Project]\Addins
HKCU\SOFTWARE\Wow6432Node\Microsoft\Office\[Word|Excel|PowerPoint|Outlook|MS Project]\Addins
HKLM\SOFTWARE\Microsoft\Office\[Word|Excel|PowerPoint|Outlook|MS Project]\Addins
HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\[Word|Excel|PowerPoint|Outlook|MS Project]\Addins

HKCU\SOFTWARE\Microsoft\Visio\Addins
HKCU\SOFTWARE\Wow6432Node\Microsoft\Visio\Addins
HKLM\SOFTWARE\Microsoft\Visio\Addins
HKLM\SOFTWARE\Wow6432Node\Microsoft\Visio\Addins

HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\[Word|Excel|PowerPoint|Outlook|MS Project]\Addins
HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Visio\Addins
HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\USER\.DEFAULT\Software\Microsoft\Office\[Word|Excel|PowerPoint|Outlook|MS Project]\Addins
HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\USER\.DEFAULT\Software\Microsoft\Visio\Addins

#LoadBehavior 
[HKCU|HKLM]\SOFTWARE\[Wow6432Node]\Microsoft\Office\[Word|Excel|PowerPoint|Outlook|MS Project]\Addins\<add-in ID>\LoadBehavior OR 
[HKCU|HKLM]\SOFTWARE\[Wow6432Node]\Microsoft\Visio]\Addins\<add-in ID>\LoadBehavior OR 
HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\[Word|Excel|PowerPoint|Outlook|MS Project]\Addins\<add-in ID>\LoadBehavior OR
HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Visio\Addins\<add-in ID>\LoadBehavior OR
HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\USER\.DEFAULT\Software\Microsoft\Office[Word|Excel|PowerPoint|Outlook|MS Project]\Addins\<add-in ID>\LoadBehavior OR
HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\USER\.DEFAULT\Software\Microsoft\Visio\Addins\<add-in ID>\LoadBehavior

#FriendlyName 
[HKCU|HKLM]\SOFTWARE\[Wow6432Node]\Microsoft\Office\[Word|Excel|PowerPoint|Outlook|MS Project]\Addins\<add-in ID>\FriendlyName OR 
[HKCU|HKLM]\SOFTWARE\[Wow6432Node]\Microsoft\Visio\Addins\<add-in ID>\FriendlyName OR 
HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\[Word|Excel|PowerPoint|Outlook|MS Project]\Addins\<add-in ID>\FriendlyName OR
HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Visio\Addins\<add-in ID>\FriendlyName OR
HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\USER\.DEFAULT\Software\Microsoft\Office\[Word|Excel|PowerPoint|Outlook|MS Project]\Addins\<add-in ID>\FriendlyName OR
HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\USER\.DEFAULT\Software\Microsoft\Visio\Addins\<add-in ID>\FriendlyName 

#Description 
[HKCU|HKLM]\SOFTWARE\[Wow6432Node]\Microsoft\Office\[Word|Excel|PowerPoint|Outlook|MS Project]\Addins\<add-in ID>\Description OR 
[HKCU|HKLM]\SOFTWARE\[Wow6432Node]\Microsoft\Visio\Addins\<add-in ID>\Description OR 
[HKCU|HKLM]\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\[Word|Excel|PowerPoint|Outlook|MS Project]\Addins\<add-in ID>\Description OR
[HKCU|HKLM]\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Visio\Addins\<add-in ID>\Description OR
HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\USER\.DEFAULT\Software\Microsoft\Office\[Word|Excel|PowerPoint|Outlook|MS Project]\Addins\<add-in ID>\ Description OR
HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\USER\.DEFAULT\Software\Microsoft\Visio\Addins\<add-in ID>\ Description 

#FullPath
#COM Add-ins
#Given the AddInId (above – it’s a ProgId) you can get the CLSID.
#The CLSID can be used to lookup the FileName in the registry at:
HKLM\SOFTWARE\Classes\[Wow6432Node]\CLSID\<CLSID>\InprocServer32\(Default) OR
HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\SOFTWARE\Classes\Wow6432Node\CLSID\<CLSID>\InprocServer32\(Default)
 
#VSTO Add-ins
#This will be defined by the Manifest key 
[HKCU|HKLM]\SOFTWARE\[Wow6432Node]\Microsoft\Office\[Word|Excel|PowerPoint|Outlook|MS Project]\Addins\<add-in ID>\Manifest OR 
[HKCU|HKLM]\SOFTWARE\[Wow6432Node]\Microsoft\Visio]\Addins\<add-in ID>\Manifest OR 
HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\[Word|Excel|PowerPoint|Outlook|MS Project]\Addins\<add-in ID>\Manifest OR
HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Visio\Addins\<add-in ID>\Manifest OR
HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\USER\.DEFAULT\Software\Microsoft\Office[Word|Excel|PowerPoint|Outlook|MS Project]\Addins\<add-in ID>\Manifest OR
HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\USER\.DEFAULT\Software\Microsoft\Visio\Addins\<add-in ID>\Manifest 

#>