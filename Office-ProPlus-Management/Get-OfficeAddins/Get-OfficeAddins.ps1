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
    
    $COMKeys = @("SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office",
                 "SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Visio\Addins",
                 "SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\USER\.DEFAULT\Software\Microsoft\Office",
                 "SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\USER\.DEFAULT\Software\Microsoft\Visio\Addins")
    
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
                        
                        $LoadBehavior = $regprov.GetDWORDValue($hkey, $addinPath, 'LoadBehavior')
                        $Description = $regProv.GetStringValue($hkey, $addinPath, 'Description')
                        $FriendlyName = $regProv.GetStringValue($hkey, $addinPath, 'FriendlyName')
    
                        $object = New-Object PSObject -Property @{Application = $officeapp; Hive = $hive; Name = $addinapp; Path = $addinpath; Type = "COM"; 
                                                                  Description = $Description.sValue; FriendlyName = $FriendlyName.sValue; LoadBehavior = $LoadBehavior.uValue}
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
                      
                        $LoadBehavior = $regprov.GetDWORDValue($hkey, $addinPath, 'LoadBehavior')
                        $Description = $regProv.GetStringValue($hkey, $addinPath, 'Description')
                        $FriendlyName = $regProv.GetStringValue($hkey, $addinPath, 'FriendlyName')
        
                        $object = New-Object PSObject -Property @{Application = $officeapp; Hive = $hive; Name = $addinapp; Path = $addinpath; Type = "VSTO"; 
                                                                  Description = $Description.sValue; FriendlyName = $FriendlyName.sValue; LoadBehavior = $LoadBehavior.uValue}
                        $object | Add-Member MemberSet PSStandardMembers $PSStandardMembers
                        $results += $object
                    }
                }
            }
        }
    }
    
    return $results;

}

function Get-CLSID {
Param(
    [string]$ComputerName = $env:COMPUTERNAME,
    [string]$ProgId
)
    $defaultDisplaySet = 'CLSID','Path','ProgID'
    $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet',[string[]]$defaultDisplaySet)
    $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
    $results = New-Object PSObject[] 0;

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
                #$value = $regProv.GetStringValue($HKLM, $progIdPath, "(Default)")
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

                    $object = New-Object PSObject -Property @{CLSID = $clsid; Path = $progIdPath; ProgID = $ProgIDValue; InprocServer32 = $InprocServer32}
                    $object | Add-Member MemberSet PSStandardMembers $PSStandardMembers
                    $results += $object
                }
            }
        }
    }

    return $results
}

function Get-ManifestKey {
Param(
    [string]$ComputerName = $env:COMPUTERNAME
)

    $regProv = Get-WmiObject -List "StdRegProv" -Namespace root\default -ComputerName $ComputerName

    $HKCU = [UInt32] "0x80000001"
    $HKLM = [UInt32] "0x80000002"

    $officeApps = @("Word","Excel","PowerPoint","Outlook","MS Project")

    $multiManifestKeys = @("SOFTWARE\Wow6432Node\Microsoft\Office",
                           "SOFTWARE\Microsoft\Office",
                           "SOFTWARE\Wow6432Node\Microsoft\Visio",
                           "SOFTWARE\Microsoft\Visio]\Addins")

    $hklmManifestKeys = @("HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office",
                          "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Visio\Addins",
                          "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\USER\.DEFAULT\Software\Microsoft\Office",
                          "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\USER\.DEFAULT\Software\Microsoft\Visio\Addins")

    

}

#AddInId 
<#
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
#>

#LoadBehavior
<#
[HKCU|HKLM]\SOFTWARE\[Wow6432Node]\Microsoft\Office\[Word|Excel|PowerPoint|Outlook|MS Project]\Addins\<add-in ID>\LoadBehavior OR 
[HKCU|HKLM]\SOFTWARE\[Wow6432Node]\Microsoft\Visio]\Addins\<add-in ID>\LoadBehavior OR 
HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\[Word|Excel|PowerPoint|Outlook|MS Project]\Addins\<add-in ID>\LoadBehavior OR
HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Visio\Addins\<add-in ID>\LoadBehavior OR
HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\USER\.DEFAULT\Software\Microsoft\Office[Word|Excel|PowerPoint|Outlook|MS Project]\Addins\<add-in ID>\LoadBehavior OR
HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\USER\.DEFAULT\Software\Microsoft\Visio\Addins\<add-in ID>\LoadBehavior
#>

#FriendlyName 
<#
[HKCU|HKLM]\SOFTWARE\[Wow6432Node]\Microsoft\Office\[Word|Excel|PowerPoint|Outlook|MS Project]\Addins\<add-in ID>\FriendlyName OR 
[HKCU|HKLM]\SOFTWARE\[Wow6432Node]\Microsoft\Visio\Addins\<add-in ID>\FriendlyName OR 
HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\[Word|Excel|PowerPoint|Outlook|MS Project]\Addins\<add-in ID>\FriendlyName OR
HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Visio\Addins\<add-in ID>\FriendlyName OR
HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\USER\.DEFAULT\Software\Microsoft\Office\[Word|Excel|PowerPoint|Outlook|MS Project]\Addins\<add-in ID>\FriendlyName OR
HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\USER\.DEFAULT\Software\Microsoft\Visio\Addins\<add-in ID>\FriendlyName 
#>

#Description
<# 
[HKCU|HKLM]\SOFTWARE\[Wow6432Node]\Microsoft\Office\[Word|Excel|PowerPoint|Outlook|MS Project]\Addins\<add-in ID>\Description OR 
[HKCU|HKLM]\SOFTWARE\[Wow6432Node]\Microsoft\Visio\Addins\<add-in ID>\Description OR 
[HKCU|HKLM]\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\[Word|Excel|PowerPoint|Outlook|MS Project]\Addins\<add-in ID>\Description OR
[HKCU|HKLM]\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Visio\Addins\<add-in ID>\Description OR
HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\USER\.DEFAULT\Software\Microsoft\Office\[Word|Excel|PowerPoint|Outlook|MS Project]\Addins\<add-in ID>\ Description OR
HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\USER\.DEFAULT\Software\Microsoft\Visio\Addins\<add-in ID>\ Description 
#>

#FullPath
##COM Add-ins
<#
#Given the AddInId (above – it’s a ProgId) you can get the CLSID.
#The CLSID can be used to lookup the FileName in the registry at:
HKLM\SOFTWARE\Classes\[Wow6432Node]\CLSID\<CLSID>\InprocServer32\(Default) OR
HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\SOFTWARE\Classes\Wow6432Node\CLSID\<CLSID>\InprocServer32\(Default)
#>

##VSTO Add-ins
<#
#This will be defined by the Manifest key 
[HKCU|HKLM]\SOFTWARE\[Wow6432Node]\Microsoft\Office\[Word|Excel|PowerPoint|Outlook|MS Project]\Addins\<add-in ID>\Manifest OR 
[HKCU|HKLM]\SOFTWARE\[Wow6432Node]\Microsoft\Visio]\Addins\<add-in ID>\Manifest OR 
HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\[Word|Excel|PowerPoint|Outlook|MS Project]\Addins\<add-in ID>\Manifest OR
HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Visio\Addins\<add-in ID>\Manifest OR
HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\USER\.DEFAULT\Software\Microsoft\Office[Word|Excel|PowerPoint|Outlook|MS Project]\Addins\<add-in ID>\Manifest OR
HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\USER\.DEFAULT\Software\Microsoft\Visio\Addins\<add-in ID>\Manifest 
#>