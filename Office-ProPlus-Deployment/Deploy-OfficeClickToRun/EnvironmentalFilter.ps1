

Function Check-UserInOU() {
   param (
      [string]$OUPath

   )

   $adSysInfo = Get-AdSystemInfo 

   $testOUPath = $OUPath + "," + $adSysInfo.DomainPath

   if ($adSysInfo.ToLower().EndsWidth($testOUPath.ToLower())) {
      return $true
   }

   return $false
}



function Get-AdSystemInfo
{
    $ADSystemInfo = New-Object -ComObject ADSystemInfo
    $adSysInfoType = $ADSystemInfo.GetType()

    $result = New-Object -TypeName PSObject -Property @{
        UserDistinguishedName = $adSysInfoType.InvokeMember('UserName','GetProperty',$null,$ADSystemInfo,$null)
        ComputerDistinguishedName = $adSysInfoType.InvokeMember('ComputerName','GetProperty',$null,$ADSystemInfo,$null)
        PDCRoleOwnerDistinguishedName = $adSysInfoType.InvokeMember('PDCRoleOwner','GetProperty',$null,$ADSystemInfo,$null)
        SchemaRoleOwnerDistinguishedName = $adSysInfoType.InvokeMember('SchemaRoleOwner','GetProperty',$null,$ADSystemInfo,$null)
        SiteName = $adSysInfoType.InvokeMember('SiteName','GetProperty',$null,$ADSystemInfo,$null)
        DomainShortName = $adSysInfoType.InvokeMember('DomainShortName','GetProperty',$null,$ADSystemInfo,$null)
        DomainDNSName = $adSysInfoType.InvokeMember('DomainDNSName','GetProperty',$null,$ADSystemInfo,$null)
        ForestDNSName = $adSysInfoType.InvokeMember('ForestDNSName','GetProperty',$null,$ADSystemInfo,$null)
        IsNativeModeDomain = $adSysInfoType.InvokeMember('IsNativeMode','GetProperty',$null,$ADSystemInfo,$null)
    }

    return $result
}