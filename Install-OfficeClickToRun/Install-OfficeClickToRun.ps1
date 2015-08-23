Add-Type -TypeDefinition @"
   public enum OfficeLanguages
   {
      CurrentOfficeLanguages,
      OSLanguage,
      OSandUserLanguages,
      AllInUseLanguages
   }
"@

Function Install-OfficeClickToRun() {

    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        [Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true, Position=0)]
        [string[]]$ComputerName = $env:COMPUTERNAME,
    
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [OfficeLanguages]$Languages = "AllInUseLanguages"
    )

    $xmlConfig = Get-CurrentOfficeConfiguration

    $officeVersion = Get-OfficeVersion -ComputerName $ComputerName
    $officeCTRs = $officeVersion | where { $_.ClickToRun -eq $true }

    if ($officeCTRs.Count -gt 0) {
       $officeCTR = $officeCTRs[0]

       $officeCTR.Version


    }
}


Function Get-CurrentOfficeConfiguration {

[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true, Position=0)]
    [string[]]$ComputerName = $env:COMPUTERNAME,
    
    [Parameter(ValueFromPipelineByPropertyName=$true)]
    [OfficeLanguages]$Languages = "AllInUseLanguages"
)

begin {
    $HKLM = [UInt32] "0x80000002"
    $HKCR = [UInt32] "0x80000000"
    $HKU = [UInt32] "0x80000003"
   
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
    [string]$versionToReport = $regProv.GetStringValue($HKLM, $configurationPath, "VersionToReport").sValue
    [string]$updatesEnabled = $regProv.GetStringValue($HKLM, $configurationPath, "UpdatesEnabled").sValue
    [string]$updateUrl = $regProv.GetStringValue($HKLM, $configurationPath, "UpdateUrl").sValue
    [string]$updateDeadline = $regProv.GetStringValue($HKLM, $configurationPath, "UpdateDeadline").sValue

    $splitProducts = $productIds.Split(',');

    if ($platform.ToLower() -eq "x86") {
        $platform = "32"
    } else {
        $platform = "64"
    }

    $osArchitecture = $os.OSArchitecture
    $osLanguage = $os.OSLanguage
    $machinelangId = "en-us"
       
    $machineCulture = [globalization.cultureinfo]::GetCultures("allCultures") | where {$_.LCID -eq $osLanguage}
    if ($machineCulture) {
        $machinelangId = $machineCulture.IetfLanguageTag
    }
    
    $primaryLanguage = checkForLanguage -langId $machinelangId

    [System.Collections.ArrayList]$additionalLanguages = New-Object System.Collections.ArrayList

    switch ($Languages) {
      "CurrentOfficeLanguages" 
      {
         $primaryLanguage = $clientCulture
      }
      "OSLanguage" 
      {
         $primaryLanguage = checkForLanguage -langId $machinelangId
      }
      "OSandUserLanguages" 
      {
         $primaryLanguage = checkForLanguage -langId $machinelangId
         $additionalLanguages = getLanguages -regProv $regProv
      }
      "AllInUseLanguages" 
      {
         $primaryLanguage = checkForLanguage -langId $machinelangId
         $additionalLanguages = getLanguages -regProv $regProv
      }
    }

    if (!($primaryLanguage)) {
        throw "Cannot find matching Office language for: $primaryLanguage"
    }
    
    foreach ($productId in $splitProducts) { 
       $excludeApps = $NULL
       if ($productId.ToLower().StartsWith("o365")) {
           $excludeApps = odtGetExcludedApps -ConfigDoc $ConfigFile -OfficeKeyPath $officeKeyPath -ProductId $productId
       }

       $officeAddLangs = odtGetOfficeLanguages -ConfigDoc $ConfigFile -OfficeKeyPath $officeKeyPath -ProductId $productId

       if (($Languages -eq "CurrentOfficeLanguages") -or ($Languages -eq "AllInUseLanguages")) {
           $additionalLanguages += $officeAddLangs
       }

       $additionalLanguages = Get-Unique -InputObject $additionalLanguages -OnType
      

       if ($additionalLanguages.Contains($primaryLanguage)) {
           $additionalLanguages.Remove($primaryLanguage)
       }

       odtAddProduct -ConfigDoc $ConfigFile -ProductId $productId -ExcludeApps $excludeApps -Version $versionToReport `
                     -Platform $platform -ClientCulture $primaryLanguage -AdditionalLanguages $additionalLanguages
       odtAddUpdates -ConfigDoc $ConfigFile -Enabled $updatesEnabled -UpdatePath $updateUrl -Deadline $updateDeadline
    }
    
    #Format-XML ([xml]($ConfigFile)) -indent 4

    return $ConfigFile
  }

  return $results;
}

}

Function Get-OfficeVersion {
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
    [string[]]$ComputerName = $env:COMPUTERNAME,
    [switch]$ShowAllInstalledProducts,
    [System.Management.Automation.PSCredential]$Credentials
)

begin {
    $HKLM = [UInt32] "0x80000002"
    $HKCR = [UInt32] "0x80000000"

    $excelKeyPath = "Excel\DefaultIcon"
    $wordKeyPath = "Word\DefaultIcon"
   
    $installKeys = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
                   'SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall'

    $officeKeys = 'SOFTWARE\Microsoft\Office',
                  'SOFTWARE\Wow6432Node\Microsoft\Office'

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

    $VersionList = New-Object -TypeName System.Collections.ArrayList
    $PathList = New-Object -TypeName System.Collections.ArrayList
    $PackageList = New-Object -TypeName System.Collections.ArrayList
    $ClickToRunPathList = New-Object -TypeName System.Collections.ArrayList
    $ConfigItemList = New-Object -TypeName System.Collections.ArrayList
    $ClickToRunList = new-object PSObject[] 0;

    foreach ($regKey in $officeKeys) {
       $officeVersion = $regProv.EnumKey($HKLM, $regKey)
       foreach ($key in $officeVersion.sNames) {
          if ($key -match "\d{2}\.\d") {
            if (!$VersionList.Contains($key)) {
              $AddItem = $VersionList.Add($key)
            }

            $path = join-path $regKey $key

            $configPath = join-path $path "Common\Config"
            $configItems = $regProv.EnumKey($HKLM, $configPath)
            foreach ($configId in $configItems.sNames) {
               $Add = $ConfigItemList.Add($configId.ToUpper())
            }

            $cltr = New-Object -TypeName PSObject
            $cltr | Add-Member -MemberType NoteProperty -Name InstallPath -Value ""
            $cltr | Add-Member -MemberType NoteProperty -Name UpdatesEnabled -Value $false
            $cltr | Add-Member -MemberType NoteProperty -Name UpdateUrl -Value ""
            $cltr | Add-Member -MemberType NoteProperty -Name StreamingFinished -Value $false
            $cltr | Add-Member -MemberType NoteProperty -Name Platform -Value ""
            $cltr | Add-Member -MemberType NoteProperty -Name ClientCulture -Value ""
            
            $packagePath = join-path $path "Common\InstalledPackages"
            $clickToRunPath = join-path $path "ClickToRun\Configuration"
            $virtualInstallPath = $regProv.GetStringValue($HKLM, $clickToRunPath, "InstallationPath").sValue

            if ($virtualInstallPath) {

            } else {
              $clickToRunPath = join-path $regKey "ClickToRun\Configuration"
              $virtualInstallPath = $regProv.GetStringValue($HKLM, $clickToRunPath, "InstallationPath").sValue
            }

            if ($virtualInstallPath) {
               if (!$ClickToRunPathList.Contains($virtualInstallPath.ToUpper())) {
                  $AddItem = $ClickToRunPathList.Add($virtualInstallPath.ToUpper())
               }

               $cltr.InstallPath = $virtualInstallPath
               $cltr.StreamingFinished = $regProv.GetStringValue($HKLM, $clickToRunPath, "StreamingFinished").sValue
               $cltr.UpdatesEnabled = $regProv.GetStringValue($HKLM, $clickToRunPath, "UpdatesEnabled").sValue
               $cltr.UpdateUrl = $regProv.GetStringValue($HKLM, $clickToRunPath, "UpdateUrl").sValue
               $cltr.Platform = $regProv.GetStringValue($HKLM, $clickToRunPath, "Platform").sValue
               $cltr.ClientCulture = $regProv.GetStringValue($HKLM, $clickToRunPath, "ClientCulture").sValue
               $ClickToRunList += $cltr
            }

            $packageItems = $regProv.EnumKey($HKLM, $packagePath)
            $officeItems = $regProv.EnumKey($HKLM, $path)

            foreach ($itemKey in $officeItems.sNames) {
              $itemPath = join-path $path $itemKey
              $installRootPath = join-path $itemPath "InstallRoot"

              $filePath = $regProv.GetStringValue($HKLM, $installRootPath, "Path").sValue
              if (!$PathList.Contains($filePath)) {
                  $AddItem = $PathList.Add($filePath)
              }
            }

            foreach ($packageGuid in $packageItems.sNames) {
              $packageItemPath = join-path $packagePath $packageGuid
              $packageName = $regProv.GetStringValue($HKLM, $packageItemPath, "").sValue
            
              if (!$PackageList.Contains($packageName)) {
                $AddItem = $PackageList.Add($packageName.Replace(' ', '').ToLower())
              }
            }

          }
       }
    }

    

    foreach ($regKey in $installKeys) {
        $keyList = new-object System.Collections.ArrayList
        $keys = $regProv.EnumKey($HKLM, $regKey)

        foreach ($key in $keys.sNames) {
           $path = join-path $regKey $key
           $installPath = $regProv.GetStringValue($HKLM, $path, "InstallLocation").sValue
           if ($installPath.Length -eq 0) { continue }

           $buildType = "64-Bit"
           if ($osArchitecture -eq "32-bit") {
              $buildType = "32-Bit"
           }

           if ($regKey.ToUpper().Contains("Wow6432Node".ToUpper())) {
              $buildType = "32-Bit"
           }

           if ($key -match "{.{8}-.{4}-.{4}-1000-0000000FF1CE}") {
              $buildType = "64-Bit" 
           }

           if ($key -match "{.{8}-.{4}-.{4}-0000-0000000FF1CE}") {
              $buildType = "32-Bit" 
           }

           if ($modifyPath) {
               if ($modifyPath.ToLower().Contains("platform=x86")) {
                  $buildType = "32-Bit"
               }

               if ($modifyPath.ToLower().Contains("platform=x64")) {
                  $buildType = "64-Bit"
               }
           }

           $primaryOfficeProduct = $false
           $officeProduct = $false
           foreach ($officeInstallPath in $PathList) {
             if ($officeInstallPath) {
                $installReg = "^" + $installPath.Replace('\', '\\')
                $installReg = $installReg.Replace('(', '\(')
                $installReg = $installReg.Replace(')', '\)')
                if ($officeInstallPath -match $installReg) { $officeProduct = $true }
             }
           }

           if (!$officeProduct) { continue };
           
           $name = $regProv.GetStringValue($HKLM, $path, "DisplayName").sValue          

           if ($ConfigItemList.Contains($key.ToUpper()) -and $name.ToUpper().Contains("MICROSOFT OFFICE")) {
              $primaryOfficeProduct = $true
           }

           $version = $regProv.GetStringValue($HKLM, $path, "DisplayVersion").sValue
           $modifyPath = $regProv.GetStringValue($HKLM, $path, "ModifyPath").sValue 

           $cltrUpdatedEnabled = $NULL
           $cltrUpdateUrl = $NULL
           $clientCulture = $NULL;

           [string]$clickToRun = $false
           if ($ClickToRunPathList.Contains($installPath.ToUpper())) {
               $clickToRun = $true
               if ($name.ToUpper().Contains("MICROSOFT OFFICE")) {
                  $primaryOfficeProduct = $true
               }

               foreach ($cltr in $ClickToRunList) {
                 if ($cltr.InstallPath) {
                   if ($cltr.InstallPath.ToUpper() -eq $installPath.ToUpper()) {
                       $cltrUpdatedEnabled = $cltr.UpdatesEnabled
                       $cltrUpdateUrl = $cltr.UpdateUrl
                       if ($cltr.Platform -eq 'x64') {
                           $buildType = "64-Bit" 
                       }
                       if ($cltr.Platform -eq 'x86') {
                           $buildType = "32-Bit" 
                       }
                       $clientCulture = $cltr.ClientCulture
                   }
                 }
               }
           }
           
           if (!$primaryOfficeProduct) {
              if (!$ShowAllInstalledProducts) {
                  continue
              }
           }

           $object = New-Object PSObject -Property @{DisplayName = $name; Version = $version; InstallPath = $installPath; ClickToRun = $clickToRun; 
                     Bitness=$buildType; ComputerName=$computer; ClickToRunUpdatesEnabled=$cltrUpdatedEnabled; ClickToRunUpdateUrl=$cltrUpdateUrl;
                     ClientCulture=$clientCulture }
           $object | Add-Member MemberSet PSStandardMembers $PSStandardMembers
           $results += $object

        }
    }

  }

  return $results;
}

}


function odtGetOfficeLanguages() {
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
    }

    process {
        $productsPath = join-path $officeKeyPath "ProductReleaseIDs\Active\$ProductId"

        $appLanguages = @() 

        $installedCultures = $regProv.EnumKey($HKLM, $productsPath)
        foreach ($installedCulture in $installedCultures.sNames) {
           if ($installedCulture.Contains("-") -and !($installedCulture.ToLower() -eq "x-none")) {
              $appLanguages += $installedCulture
           }
        }
        
        return $appLanguages;
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
              $appsToExclude += $appName
           }
        }
        
        return $appsToExclude;
    }
}

function odtAddProduct() {
    param(
       [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
       [System.XML.XMLDocument]$ConfigDoc = $NULL,

       [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
       [string]$ProductId = $NULL,

       [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
       [string]$Platform = $NULL,

       [Parameter(ValueFromPipelineByPropertyName=$true)]
       [string]$ClientCulture = "en-us",

       [Parameter(ValueFromPipelineByPropertyName=$true)]
       [string[]]$AdditionalLanguages,

       [Parameter(ValueFromPipelineByPropertyName=$true)]
       [string[]] $ExcludeApps,

       [Parameter(ValueFromPipelineByPropertyName=$true)]
       [string]$Version = $NULL

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

    if ($Version) {
       $AddElement.SetAttribute("Version", $Version) | Out-Null
    }

    if ($Platform) {
       $AddElement.SetAttribute("Edition", $Platform) | Out-Null
    }

    [System.XML.XMLElement]$ProductElement = $ConfigFile.Configuration.Add.Product | ?  ID -eq $ProductId
    if($ProductElement -eq $null){
        [System.XML.XMLElement]$ProductElement=$ConfigFile.CreateElement("Product")
        $AddElement.appendChild($ProductElement) | Out-Null
        $ProductElement.SetAttribute("ID", $ProductId) | Out-Null
    }

    $LanguageIds = @($ClientCulture)

    foreach ($addLang in $AdditionalLanguages) {
       $LanguageIds += $addLang 
    }

    foreach($LanguageId in $LanguageIds){
        [System.XML.XMLElement]$LanguageElement = $ProductElement.Language | ?  ID -eq $LanguageId
        if($LanguageElement -eq $null){
            [System.XML.XMLElement]$LanguageElement=$ConfigFile.CreateElement("Language")
            $ProductElement.appendChild($LanguageElement) | Out-Null
            $LanguageElement.SetAttribute("ID", $LanguageId.ToString().ToLower()) | Out-Null
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

function odtAddUpdates{

    [CmdletBinding()]
    Param(

        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [System.XML.XMLDocument]$ConfigDoc = $NULL,
        
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $Enabled,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $UpdatePath,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $TargetVersion,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $Deadline

    )

    Process{
        #Check to make sure the correct root element exists
        if($ConfigDoc.Configuration -eq $null){
            throw $NoConfigurationElement
        }

        #Get the Updates Element if it exists
        [System.XML.XMLElement]$UpdateElement = $ConfigDoc.Configuration.GetElementsByTagName("Updates").Item(0)
        if($ConfigDoc.Configuration.Updates -eq $null){
            [System.XML.XMLElement]$UpdateElement=$ConfigDoc.CreateElement("Updates")
            $ConfigDoc.Configuration.appendChild($UpdateElement) | Out-Null
        }

        #Set the desired values
        if([string]::IsNullOrWhiteSpace($Enabled) -eq $false){
            $UpdateElement.SetAttribute("Enabled", $Enabled) | Out-Null
        } else {
          if ($PSBoundParameters.ContainsKey('Enabled')) {
              $ConfigDoc.Configuration.Updates.RemoveAttribute("Enabled")
          }
        }

        if([string]::IsNullOrWhiteSpace($UpdatePath) -eq $false){
            $UpdateElement.SetAttribute("UpdatePath", $UpdatePath) | Out-Null
        } else {
          if ($PSBoundParameters.ContainsKey('UpdatePath')) {
              $ConfigDoc.Configuration.Updates.RemoveAttribute("UpdatePath")
          }
        }

        if([string]::IsNullOrWhiteSpace($TargetVersion) -eq $false){
            $UpdateElement.SetAttribute("TargetVersion", $TargetVersion) | Out-Null
        } else {
          if ($PSBoundParameters.ContainsKey('TargetVersion')) {
              $ConfigDoc.Configuration.Updates.RemoveAttribute("TargetVersion")
          }
        }

        if([string]::IsNullOrWhiteSpace($Deadline) -eq $false){
            $UpdateElement.SetAttribute("Deadline", $Deadline) | Out-Null
        } else {
          if ($PSBoundParameters.ContainsKey('Deadline')) {
              $ConfigDoc.Configuration.Updates.RemoveAttribute("Deadline")
          }
        }

    }
}

function Format-XML ([xml]$xml, $indent=2) { 
    $StringWriter = New-Object System.IO.StringWriter 
    $XmlWriter = New-Object System.XMl.XmlTextWriter $StringWriter 
    $xmlWriter.Formatting = "indented" 
    $xmlWriter.Indentation = $Indent 
    $xml.WriteContentTo($XmlWriter) 
    $XmlWriter.Flush() 
    $StringWriter.Flush() 
    Write-Output $StringWriter.ToString() 
}

function getLanguages() {
    param(
       [Parameter(ValueFromPipelineByPropertyName=$true)]
       $regProv = $NULL
    )

  [System.Collections.ArrayList] $returnLangs = New-Object System.Collections.ArrayList

  $HKU = [UInt32] "0x80000003"
  $userKeys = $regProv.EnumKey($HKU, "");

  foreach ($userKey in $userKeys.sNames) {
     if ($userKey.Length -gt 8 -and !($userKey.ToLower().EndsWith("_classes"))) {
       [string]$userProfilePath = join-path $userKey "Control Panel\International\User Profile"
       [string[]]$userLanguages = $regProv.GetMultiStringValue($HKU, $userProfilePath, "Languages").sValue
       foreach ($userLang in $userLanguages) {
         $convertLang = checkForLanguage -langId $userLang 
         if ($convertLang) {
             $returnLangs += $convertLang.ToLower()
         }
       }
        
     }
  }

  $returnLangs = $returnLangs | Get-Unique 
  return $returnLangs

}

function checkForLanguage() {
    param(
       [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
       [string]$langId = $NULL
    )

    if ($availableLangs.Contains($langId.ToLower())) {
       return $langId
    } else {
       $langStart = $langId.Split('-')[0]
       $checkLang = $NULL

       foreach ($availabeLang in $availableLangs) {
          if ($availabeLang.ToLower().StartsWith($langStart.ToLower())) {
             $checkLang = $availabeLang
             break;
          }
       }

       return $checkLang
    }
}

$availableLangs = @("en-us",
"ar-sa","bg-bg","zh-cn","zh-tw","hr-hr","cs-cz","da-dk","nl-nl","et-ee",
"fi-fi","fr-fr","de-de","el-gr","he-il","hi-in","hu-hu","id-id","it-it",
"ja-jp","kk-kh","ko-kr","lv-lv","lt-lt","ms-my","nb-no","pl-pl","pt-br",
"pt-pt","ro-ro","ru-ru","sr-latn-rs","sk-sk","sl-si","es-es","sv-se","th-th",
"tr-tr","uk-ua");

Install-OfficeClickToRun

