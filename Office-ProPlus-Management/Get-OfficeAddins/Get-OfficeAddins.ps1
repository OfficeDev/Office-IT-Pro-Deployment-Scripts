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
                "Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Visio",
                "Microsoft\Office\ClickToRun\REGISTRY\USER\.DEFAULT\Software\Microsoft\Visio")

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