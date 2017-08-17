#Create a WMI Class
function New-OfficeAddinWMIClass{
Param(
    [Paraneter()]
    [string]$ClassName = "Custom_OfficeAddins",

    [Paraneter()]
    [string]$NameSpace = "root\cimv2"
)
    $NewClass = New-Object System.Management.ManagementClass($NameSpace, $null, $null)
    $NewClass.Name = $ClassName
    $NewClass.Put() | Out-Null
}

#Create properties for the new WMI Class
Function New-WMIClassProperty {
[CmdletBinding()]
Param(
	[Parameter()]
      [string]$ClassName = "Custom_OfficeAddins",

      [Parameter()]
      [string]$NameSpace = "Root\cimv2",

      [Parameter()]
      [string]$PropertyName,

      [Parameter()]
      [string]$PropertyValue
)
   
    [WmiClass]$WmiClass = Get-WmiObject -Class $ClassName -Namespace $NameSpace -List        
    $WmiClass.Properties.Add($PropertyName,$PropertyValue)          
    $WmiClass.Put() | Out-Null
}

#Set the Class properties
Function Set-WMIPropertyValue {
[CmdletBinding()]
Param(
	[Parameter()]
       [string]$ClassName = "Custom_OfficeAddins",

       [Parameter()]
       [string]$NameSpace = "Root\cimv2",

       [Parameter()]
       [string]$PropertyName,

       [Parameter()]
       [string]$PropertyValue
)
    
    [WmiClass]$WmiClass = Get-WmiObject -Class $ClassName -Namespace $NameSpace -List       
    $WMI_Class.SetPropertyValue($PropertyName,$PropertyValue)          
    $WMI_Class.Put() | Out-Null  
}

#Set the Class Property Qualifier
Function Set-WMIPropertyQualifier {
[CmdletBinding()]
Param(
	[Parameter()]
       [string]$ClassName = "Custom_OfficeAddins",

       [Parameter()]
       [string]$NameSpace="Root\cimv2",

       [Parameter()]
       [string]$PropertyName,

       [Parameter()]
       [string]$QualifierName,

       [Parameter()]
       [string]$QualifierValue,

       [switch]$key,

       [switch]$IsAmended = $false,

       [switch]$IsLocal = $true,

       [switch]$PropagatesToInstance = $true,

       [switch]$PropagesToSubClass = $false,

       [switch]$IsOverridable = $true
)
  
    [WmiClass]$WmiClass = Get-WmiObject -Class $ClassName -Namespace $NameSpace -List  

    if ($WmiClass.Properties[$PropertyName]){
        if($Key){
            $WmiClass.Properties[$PropertyName].Qualifiers.Add("Key",$true)
            $WmiClass.put() | Out-Null
        }else{ 
            $WmiClass.Properties[$PropertyName].Qualifiers.Add($QualifierName,$QualifierValue,$IsAmended,$IsLocal,$PropagatesToInstance,$PropagesToSubClass)
            $WmiClass.put() | Out-Null
        }
    }
}

#Create a new Class Instance
Function New-WMIOfficeAddinClassInstance {
[CmdletBinding()]
Param(
	[Parameter()]
       [string]$ClassName = "Custom_OfficeAddins",

       [Parameter()]
       [string]$NameSpace = "Root\cimv2",

       [switch]$PutInstance
)
    
    [WmiClass]$WmiClass = Get-WmiObject -Class $ClassName -Namespace $NameSpace -List 
              
    if($PutInstance){   
        $PutInstance.Put()
    }else{
        $WmiClass.CreateInstance()
    }
}

$MyNewInstance = New-WMIOfficeAddinClassInstance -ClassName $ClassName
$MyNewInstance.Application = $Application
$MyNewInstance.ComputerName = $env:COMPUTERNAME
$MyNewInstance.Description = $Description
$MyNewInstance.FriendlyName = $FriendlyName
$MyNewInstance.FullPath = $FullPath
$MyNewInstance.LoadBehavior = $LoadBehavior
$MyNewInstance.LoadTime = $LoadTime
$MyNewInstance.Name = $Name
$MyNewInstance.OfficeVersion = $OfficeVersion
$MyNewInstance.RegistryPath = $RegistryPath

New-WMIOfficeAddinClassInstance -ClassName $ClassName -PutInstance $MyNewInstance