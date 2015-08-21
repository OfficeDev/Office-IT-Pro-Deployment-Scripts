

Function Get-CurrentOfficeConfiguration {
<#
.Synopsis
Gets the Office Version installed on the computer

.DESCRIPTION
This function will query the local or a remote computer and return the information about Office Products installed on the computer

.NOTES   
Name: Get-OfficeVersion
Version: 1.0.3
DateCreated: 2015-07-01
DateUpdated: 2015-07-21

.LINK
https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts

.PARAMETER ComputerName
The computer or list of computers from which to query 

.PARAMETER ShowAllInstalledProducts
Will expand the output to include all installed Office products

.EXAMPLE
Get-OfficeVersion

Description:
Will return the locally installed Office product

.EXAMPLE
Get-OfficeVersion -ComputerName client01,client02

Description:
Will return the installed Office product on the remote computers

.EXAMPLE
Get-OfficeVersion | select *

Description:
Will return the locally installed Office product with all of the available properties

#>
[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true, Position=0)]
    [string[]]$ComputerName = $env:COMPUTERNAME
)

begin {
    $HKLM = [UInt32] "0x80000002"
    $HKCR = [UInt32] "0x80000000"
   
    $installKeys = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
                   'SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall'

    $officeKeys = 'SOFTWARE\Microsoft\Office\15.0\ClickToRun',
                  'SOFTWARE\Wow6432Node\Microsoft\Office\15.0\ClickToRun'

    $defaultDisplaySet = 'DisplayName','Version', 'ComputerName'

    $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet(‘DefaultDisplayPropertySet’,[string[]]$defaultDisplaySet)
    $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
}


process {

 $results = new-object PSObject[] 0;

 foreach ($computer in $ComputerName) {
    if ($Credentials) {
       $os=Get-WMIObject win32_operatingsystem -computername $computer -Credential $Credentials
    } else {
       $os=Get-WMIObject win32_operatingsystem -computername $computer
    }

    $osArchitecture = $os.OSArchitecture

    if ($Credentials) {
       $regProv = Get-Wmiobject -list "StdRegProv" -namespace root\default -computername $computer -Credential $Credentials
    } else {
       $regProv = Get-Wmiobject -list "StdRegProv" -namespace root\default -computername $computer
    }

    [System.XML.XMLDocument]$ConfigFile = New-Object System.XML.XMLDocument

    [string]$officeKeyPath = "";

    foreach ($regPath in $officeKeys) {
       [string]$installPath = $regProv.GetStringValue($HKLM, $regPath, "InstallPath").sValue
       if ($installPath) {
          if ($installPath.Length -gt 0) {
              $officeKeyPath = $regPath;
              break;
          }
       }
    }

    $configurationPath = join-path $officeKeyPath "Configuration"

    [string]$platform = $regProv.GetStringValue($HKLM, $configurationPath, "Platform").sValue
    [string]$clientCulture = $regProv.GetStringValue($HKLM, $configurationPath, "ClientCulture").sValue
    [string]$productIds = $regProv.GetStringValue($HKLM, $configurationPath, "ProductReleaseIds").sValue
    
    $productIds = $productIds.Split(',');

    foreach ($productId in $productIds) {       
       $excludeApps = odtGetExcludedApps -ConfigDoc $ConfigFile -OfficeKeyPath $officeKeyPath -ProductId $productId
       odtAddProduct -ConfigDoc $ConfigFile -ProductId $productId -ExcludeApps $excludeApps
    }
    
    
    
    $ConfigFile.OuterXml
  }

  return $results;
}

}


function odtGetExcludedApps() {
    param(
       [Parameter(ValueFromPipelineByPropertyName=$true)]
       [System.XML.XMLDocument]$ConfigDoc = $NULL,
              
       [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
       [string]$OfficeKeyPath = $NULL,

       [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
       [string]$ProductId = $NULL
    )

    begin {
        $HKLM = [UInt32] "0x80000002"
        $HKCR = [UInt32] "0x80000000"

        $allExcludeApps = 'Access','Excel','Groove','InfoPath','Lync','OneNote','Outlook',
                       'PowerPoint','Publisher','Word'
        #"SharePointDesigner","Visio", 'Project'
    }

    process {
        $productsPath = join-path $officeKeyPath "ProductReleaseIDs\Active\$ProductId\x-none"

        $appsToExclude = @() 

        $installedItems = $regProv.EnumKey($HKLM, $productsPath)
        foreach ($appName in $allExcludeApps) {
           [bool]$appInstalled = $false
           foreach ($installedItem in $installedItems.sNames) {
               if ($installedItem.ToLower().StartsWith($appName.ToLower())) {
                  $appInstalled = $true
                  break;
              }
           }
           
           if (!($appInstalled)) {
              $appsToExclude += $excludeApp
           }
        }
        
        return $appsToExclude;
    }
}

function odtAddProduct() {
    param(
       [Parameter(ValueFromPipelineByPropertyName=$true)]
       [System.XML.XMLDocument]$ConfigDoc = $NULL,

       [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
       [string]$ProductId = $NULL,

       [Parameter(ValueFromPipelineByPropertyName=$true)]
       [string]$ClientCulture = "en-us",

       [Parameter(ValueFromPipelineByPropertyName=$true)]
       [string[]]$AdditionalLanguages,

       [Parameter(ValueFromPipelineByPropertyName=$true)]
       [string[]] $ExcludeApps
    )

    [System.XML.XMLElement]$ConfigElement=$NULL
    if($ConfigFile.Configuration -eq $null){
        $ConfigElement=$ConfigFile.CreateElement("Configuration")
        $ConfigFile.appendChild($ConfigElement) | Out-Null
    }

    [System.XML.XMLElement]$AddElement=$NULL
    if($ConfigFile.Configuration.Add -eq $null){
        $AddElement=$ConfigFile.CreateElement("Add")
        $ConfigFile.DocumentElement.appendChild($AddElement) | Out-Null
    } else {
        $AddElement = $ConfigFile.Configuration.Add 
    }

    [System.XML.XMLElement]$ProductElement = $ConfigFile.Configuration.Add.Product | ?  ID -eq $ProductId
    if($ProductElement -eq $null){
        [System.XML.XMLElement]$ProductElement=$ConfigFile.CreateElement("Product")
        $AddElement.appendChild($ProductElement) | Out-Null
        $ProductElement.SetAttribute("ID", $ProductId) | Out-Null
    }

    $LanguageIds = @($ClientCulture)

    foreach ($addLang in $AdditionalLanguages) {
       $LanguageIds.Add($addLang)
    }

    foreach($LanguageId in $LanguageIds){
        [System.XML.XMLElement]$LanguageElement = $ProductElement.Language | ?  ID -eq $LanguageId
        if($LanguageElement -eq $null){
            [System.XML.XMLElement]$LanguageElement=$ConfigFile.CreateElement("Language")
            $ProductElement.appendChild($LanguageElement) | Out-Null
            $LanguageElement.SetAttribute("ID", $LanguageId) | Out-Null
        }
    }

    foreach($ExcludeApp in $ExcludeApps){
        [System.XML.XMLElement]$ExcludeAppElement = $ProductElement.ExcludeApp | ?  ID -eq $ExcludeApp
        if($ExcludeAppElement -eq $null){
            [System.XML.XMLElement]$ExcludeAppElement=$ConfigFile.CreateElement("ExcludeApp")
            $ProductElement.appendChild($ExcludeAppElement) | Out-Null
            $ExcludeAppElement.SetAttribute("ID", $ExcludeApp) | Out-Null
        }
    }

}

Get-CurrentOfficeConfiguration
