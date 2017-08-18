[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [Parameter()]
    [string]$ClassName = "Custom_OfficeAddins2",

    [Parameter()]
    [string]$NameSpace = "root\cimv2"
)

#Create a WMI Class
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

Function Get-WMIClass{
[CmdletBinding()]
	Param(
		[Parameter(Mandatory=$false,valueFromPipeLine=$true)][string]$ClassName,
        [Parameter(Mandatory=$false)][string]$NameSpace = "root\cimv2"
	
	)  
    begin{
    write-verbose "Getting WMI class $($Classname)"
    }
    Process{
        if (!($ClassName)){
            $return = Get-WmiObject -Namespace $NameSpace -Class * -list
        }else{
            $return = Get-WmiObject -Namespace $NameSpace -Class $ClassName -list
        }
    }
    end{

        return $return
    }
}

#Create properties for the new WMI Class
Function New-OfficeAddinsWMIClassProperty {
[CmdletBinding()]
Param(
	[Parameter()]
      [string]$ClassName = "Custom_OfficeAddins",

      [Parameter()]
      [string]$NameSpace = "Root\cimv2",

      [Parameter()]
      [string]$PropertyName,

      [Parameter()]
      [string]$PropertyValue = ""
)   

    [WmiClass]$WmiClass = Get-WmiObject -Class $ClassName -Namespace $NameSpace -List -ErrorAction SilentlyContinue
            
    $WmiClass.Properties.Add($PropertyName,$PropertyValue)
                
    $WmiClass.Put() | Out-Null
}

<#
#foreach($property in $PropertyName){
#    switch($property){
#        "Application" {
#            New-WMIClassProperty -PropertyName $property -PropertyValue $null
#        }
#        "ComputerName"{
#            New-WMIClassProperty -PropertyName $property -PropertyValue $null
#        }
#        "Description"{
#            New-WMIClassProperty -PropertyName $property -PropertyValue $null
#        }
#        "FriendlyName"{
#            New-WMIClassProperty -PropertyName $property -PropertyValue $null
#        }
#        "FullPath"{
#            New-WMIClassProperty -PropertyName $property -PropertyValue $null
#        }
#        "LoadBehavior"{
#            New-WMIClassProperty -PropertyName $property -PropertyValue $null
#        }
#        "LoadTime"{
#            New-WMIClassProperty -PropertyName $property -PropertyValue $null
#        }
#        "Name"{
#            New-WMIClassProperty -PropertyName $property -PropertyValue $null
#        }
#        "OfficeVersion"{
#            New-WMIClassProperty -PropertyName $property -PropertyValue $null
#        }
#        "RegistryPath"{
#            New-WMIClassProperty -PropertyName $property -PropertyValue $null                    
#        }
#    }
#}
#>

#Set the Class properties
<#Function Set-WMIPropertyValue {
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
    
    [WmiClass]$WmiClass = Get-WmiObject -Class $ClassName -Namespace $NameSpace -List -ErrorAction SilentlyContinue       
    $WMI_Class.SetPropertyValue($PropertyName,$PropertyValue)          
    $WMI_Class.Put() | Out-Null  
}
#>

#Set the Class Property Qualifier
Function Set-OfficeAddinWMIPropertyQualifier {
[CmdletBinding()]
Param(
	[Parameter()]
       [string]$ClassName = "Custom_OfficeAddins",

       [Parameter()]
       [string]$NameSpace = "Root\cimv2",

       [Parameter()]
       [string]$PropertyName,

       [Parameter()]
       [string]$QualifierName = "Key",

       [Parameter()]
       [string]$QualifierValue = $true,

       [switch]$key,

       [switch]$IsAmended = $false,

       [switch]$IsLocal = $true,

       [switch]$PropagatesToInstance = $true,

       [switch]$PropagesToSubClass = $false,

       [switch]$IsOverridable = $true
)
  
    #[WmiClass]$WmiClass = Get-WmiObject -Class $ClassName -Namespace $NameSpace -List -ErrorAction SilentlyContinue 

    [WmiClass]$WmiClass = Get-WMIClass -ClassName $ClassName -NameSpace $NameSpace

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

       [Parameter(valueFromPipeLine=$true)]$PutInstance
)
    
    [WmiClass]$WmiClass = Get-WmiObject -Class $ClassName -Namespace $NameSpace -List -ErrorAction SilentlyContinue 
              
    if($PutInstance){   
        $PutInstance.Put()
    }else{
        $WmiClass.CreateInstance()
    }
}

New-OfficeAddinWMIClass

New-OfficeAddinsWMIClassProperties

Set-WMIPropertyQualifier -ClassName Custom_OfficeAddins -PropertyName Name -QualifierName Key -QualifierValue $true

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

#$Application   = "MS Project"
#$ComputerName  = "VCG-ADAMS"
#$Description   = "Team Foundation Add-in"
#$FriendlyName  = "Team Foundation Add-in"
#$FullPath      = "C:\Program Files\Common Files\Microsoft Shared\Team Foundation Server\14.0\x64\TFSOfficeAdd-in.dll"
#$LoadBehavior  = "0"
#$LoadTime      = "0"
#$Name          = "TFCOfficeShim.Connect.14"
#$OfficeVersion = "0"
#$RegistryPath  = "Software\Microsoft\Office\MS Project\Addins\TFCOfficeShim.Connect.14"
#
#$Application   = "MS Outlook"
#$ComputerName  = "VCG-ADAMS"
#$Description   = "Team Foundation Add-in"
#$FriendlyName  = "Team Foundation Add-in"
#$FullPath      = "C:\Program Files\Common Files\Microsoft Shared\Team Foundation Server\14.0\x64\TFSOfficeAdd-in.dll"
#$LoadBehavior  = "0"
#$LoadTime      = "0"
#$Name          = "Outlook.Addin"
#$OfficeVersion = "0"
#$RegistryPath  = "Software\Microsoft\Office\MS Project\Addins\TFCOfficeShim.Connect.14"
#
#"Application", "ComputerName", "Description", "FriendlyName", "FullPath", "LoadBehavior", "LoadTime","Name", "OfficeVersion", "RegistryPath"  