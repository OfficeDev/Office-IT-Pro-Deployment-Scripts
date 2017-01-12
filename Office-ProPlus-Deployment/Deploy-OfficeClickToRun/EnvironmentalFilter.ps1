Function Check-UserInOUPath() {
   param (
      [parameter(Mandatory=$true)]
      [string]$ContainerPath,

      [parameter()]
      [bool]$IncludeSubContainers=$true
   )

   $adSysInfo = Get-AdSystemInfo
   return checkInOUPath -DistinguishedName $adSysInfo.UserDistinguishedName -IncludeSubContainers $IncludeSubContainers -ContainerPath $ContainerPath
}

Function Check-ComputerInOUPath() {
   param (
      [parameter(Mandatory=$true)]
      [string]$ContainerPath,

      [parameter()]
      [bool]$IncludeSubContainers=$true
   )


   $adSysInfo = Get-AdSystemInfo
   return checkInOUPath -DistinguishedName $adSysInfo.ComputerDistinguishedName -IncludeSubContainers $IncludeSubContainers -ContainerPath $ContainerPath
}


Function checkInOUPath() {
   param (
      [parameter(Mandatory=$true)]
      [string]$DistinguishedName,

      [parameter(Mandatory=$true)]
      [string]$ContainerPath,

      [parameter()]
      [bool]$IncludeSubContainers=$true
   )

   $OUPath = $OUPath -replace ",$", ""

   $pathParse = Parse-LDAPPath -DistinguishedName $DistinguishedName
   
   if ($ContainerPath.ToUpper().Contains("DC=")) {
     $chkPath = $pathParse.ContainerPath.ToUpper() + "," + $pathParse.DomainPath

     if ($IncludeSubContainers) {
         if ($chkPath.ToUpper().EndsWith($ContainerPath.ToUpper())) {
            return $true
         }
     } else {
         if ($ContainerPath.ToUpper() -eq $chkPath.ToUpper()) {
            return $true
         }
     }
   } else {
     if ($IncludeSubContainers) {
         if ($pathParse.ContainerPath.ToUpper().Trim().EndsWith($ContainerPath.ToUpper().Trim())) {
            return $true
         }
     } else {
         if ($ContainerPath.ToUpper().Trim() -eq $pathParse.ContainerPath.ToUpper().Trim()) {
            return $true
         }
     }
   }
   return $false
}

function Parse-LDAPPath() {
   param(
      [Parameter(mandatory=$true)]
      [string]$DistinguishedName
   )

   $userDN = $DistinguishedName.Replace("\,", "-----")

   $pathSplit = $userDN.Split(',')
   $commonName = $pathSplit[0]

   $contPath = "";
   $domainPath = "";
   for ($n=1;$n -lt $pathSplit.Length;$n++) {
     $pathItem = $pathSplit[$n]
      
     if ($pathItem.ToUpper().StartsWith("OU") -or $pathItem.ToUpper().StartsWith("CN")) {
        if ($contPath.Length -gt 0) { $contPath += "," }
        $contPath += $pathItem 
     }

     if ($pathItem.ToUpper().StartsWith("DC")) {
        if ($domainPath.Length -gt 0) { $domainPath += "," }
        $domainPath += $pathItem 
     }
   }

    $contPath = $contPath -replace ",$", ""

    $result = New-Object -TypeName PSObject -Property @{
        CommonName = $commonName.Replace("-----", "\,")
        ContainerPath = $contPath.Replace("-----", "\,")
        DomainPath = $domainPath.Replace("-----", "\,")
    }  
    return $result
}

function Get-AdSystemInfo {
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
