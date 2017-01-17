﻿try {
Add-Type -ErrorAction SilentlyContinue -TypeDefinition @"
   public enum OfficeLanguages
   {
      CurrentOfficeLanguages,
      OSLanguage,
      OSandUserLanguages,
      AllInUseLanguages
   }
"@
} catch {}

Function Generate-ODTConfigurationXml {
<#
.Synopsis
Generates the Office Deployment Tool (ODT) Configuration XML from the current configuration of the target computer
.DESCRIPTION
This function will query the local or a remote computer and Generate the ODT configuration xml based on the local Office install
and the local languages that are used on the local computer.  If Office isn't installed then it will utilize the configuration file
specified in the 
.NOTES   
Name: Generate-ODTConfigurationXml
Version: 1.0.3
DateCreated: 2015-08-24
DateUpdated: 2016-06-13
.LINK
https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts
.PARAMETER ComputerName
The computer or list of computers from which to query 
.PARAMETER Languages
Will expand the output to include all installed Office products
.PARAMETER TargetFilePath
The path and file name of the file to save the Configuration xml
.PARAMETER IncludeUpdatePathAsSourcePath
If this parameter is set to $true then the SourcePath in the Configuration xml will be set to 
the current UpdatePath on the local computer.  This assumes that the UpdatePath location has 
the required files needed to run the installation 
.PARAMETER DefaultConfigurationXml
This parameter sets the path to the Default Configuration XML file.  If Office is not installed on
the computer that this script is run against it will default to this file in order to generate the 
ODT Configuration XML.  The default file should have the products that you would want installed on 
a workstation if Office isn't currently installed.  If this parameter is set to $NULL then it will
not generate configuration XML if Office is not installed.  By default the script looks for a file
called "DefaultConfiguration.xml" in the same directory as the script
.EXAMPLE
Generate-ODTConfigurationXml | fl
Description:
Will generate the Office Deployment Tool (ODT) configuration XML based on the local computer
.EXAMPLE
Generate-ODTConfigurationXml  -ComputerName client01,client02 | fl
Description:
Will generate the Office Deployment Tool (ODT) configuration XML based on the configuration of the remote computers client01 and client02
.EXAMPLE
Generate-ODTConfigurationXml -Languages OSandUserLanguages
Description:
Will generate the Office Deployment Tool (ODT) configuration XML based on the local computer and add the languages that the Operating System and the local users
are currently using.
.EXAMPLE
Generate-ODTConfigurationXml -Languages OSLanguage
Description:
Will generate the Office Deployment Tool (ODT) configuration XML based on the local computer and add the Current UI Culture language of the Operating System
.EXAMPLE
Generate-ODTConfigurationXml -Languages CurrentOfficeLanguages
Description:
Will generate the Office Deployment Tool (ODT) configuration XML based on the local computer and add only add the Languages currently in use by the current Office installation
#>

[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [Parameter(ValueFromPipelineByPropertyName=$true, Position=0)]
    [string[]]$ComputerName = $env:COMPUTERNAME,
    
    [Parameter(ValueFromPipelineByPropertyName=$true)]
    [OfficeLanguages]$Languages = "AllInUseLanguages",

    [Parameter(ValueFromPipelineByPropertyName=$true)]
    [String]$TargetFilePath = $NULL,

    [Parameter(ValueFromPipelineByPropertyName=$true)]
    [bool]$IncludeUpdatePathAsSourcePath = $false,

    [Parameter(ValueFromPipelineByPropertyName=$true)]
    [string]$DefaultConfigurationXml = $NULL
)

begin {
    $HKLM = [UInt32] "0x80000002"
    $HKCR = [UInt32] "0x80000000"
    $HKU = [UInt32] "0x80000003"
   
    $installKeys = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
                   'SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall'

    $defaultDisplaySet = 'DisplayName','Version', 'ComputerName'

    $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet',[string[]]$defaultDisplaySet)
    $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
    
    [string]$tempStr = $MyInvocation.MyCommand.Path

    $scriptPath = GetScriptPath

    if (!($DefaultConfigurationXml)) {
      $DefaultConfigurationXml = (Join-Path $scriptPath "DefaultConfiguration.xml") 
    }
}

process {
    # write log
    $lineNum = Get-CurrentLineNumber    
    $filName = Get-CurrentFileName 
    WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "begin function"

 if ($TargetFilePath) {
     $folderPath = Split-Path -Path $TargetFilePath -Parent
     $fileName = Split-Path -Path $TargetFilePath -Leaf
     if ($folderPath) {
         [system.io.directory]::CreateDirectory($folderPath) | Out-Null
     }
 }
 
 $results = new-object PSObject[] 0;

 foreach ($computer in $ComputerName) {
   try {
    if ($Credentials) {
       $os=Get-WMIObject win32_operatingsystem -computername $computer -Credential $Credentials -ErrorAction Stop
    } else {
       $os=Get-WMIObject win32_operatingsystem -computername $computer  -ErrorAction Stop
    }

    if ($Credentials) {
       $regProv = Get-Wmiobject -list "StdRegProv" -namespace root\default -computername $computer -Credential $Credentials  -ErrorAction Stop
    } else {
       $regProv = Get-Wmiobject -list "StdRegProv" -namespace root\default -computername $computer  -ErrorAction Stop
    }

    if ($TargetFilePath) {
      if ($ComputerName.Length -gt 1) {
         $NewFileName = $computer + "-" + $fileName
         $TargetFilePath = Join-Path $folderPath $NewFileName
      }
    }

    [System.XML.XMLDocument]$ConfigFile = New-Object System.XML.XMLDocument

    $productReleaseIds = "";
    $productPlatform = "32";

    $officeConfig = getCTRConfig -regProv $regProv
    $mainOfficeProduct = Get-OfficeVersion -ComputerName $ComputerName
    $officeProducts = Get-OfficeVersion -ComputerName $ComputerName -ShowAllInstalledProducts

    if (!($officeConfig.ClickToRunInstalled)) {
        $officeConfig = getOfficeConfig -regProv $regProv -mainOfficeProduct $mainOfficeProduct -officeProducts $officeProducts

        if ($officeConfig -and $officeConfig.OfficeKeyPath) {
            $officeLangs = officeGetLanguages -regProv $regProv -OfficeKeyPath $officeConfig.OfficeKeyPath
        }
        if ($officeConfig -and $officeConfig.Platform) {
           $productPlatform = $officeConfig.Platform
        }
    } else {
      $productPlatform = $officeConfig.Platform
      $otherProducts = $officeConfig.ProductReleaseIds
      $otherProducts = generateProductReleaseIds -OfficeProducts $officeProducts -MainOfficeProduct $mainOfficeProduct
    }

    if ($officeConfig.ProductReleaseIds) {
        $productReleaseIds = $officeConfig.ProductReleaseIds
    }

    if ($otherProducts) {
        if ($productReleaseIds) {
            $productReleaseIds += ",$otherProducts"
        } else {
            $productReleaseIds += $otherProducts
        }
    }

    [bool]$officeExists = $true

    if (!($officeProducts)) {
      $officeExists = $false
      if ($DefaultConfigurationXml) {
          if (Test-Path -Path $DefaultConfigurationXml) {
             $ConfigFile.Load($DefaultConfigurationXml)

             $products = $ConfigFile.SelectNodes("/Configuration/Add/Product")
             if ($products) {
                 foreach ($product in $products) {
                    if ($productReleaseIds.Length -gt 0) { $productReleaseIds += "," }
                    $productReleaseIds += $product.ID
                 }
             }

             $addNode = $ConfigFile.SelectSingleNode("/Configuration/Add");
             if ($addNode) {
                $productPlatform = $addNode.OfficeClientEdition
             }

          }
      }
    }
    
    if ($productReleaseIds) {
        $splitProducts = $productReleaseIds.Split(',');
        
        $newSplitProducts = @()
        foreach ($productId in $splitProducts) {
            if($productId.ToUpper() -notlike "SPD*") {     
                if (!($newSplitProducts -Contains $productId)) {
                   $newSplitProducts += $productId
                }
            }
        }
        
        $splitProducts = $newSplitProducts
    }

    $osArchitecture = $os.OSArchitecture
    $osLanguage = $os.OSLanguage
    $machinelangId = "en-us"
       
    $machineCulture = [globalization.cultureinfo]::GetCultures("allCultures") | where {$_.LCID -eq $osLanguage}
    if ($machineCulture) {
        $machinelangId = $machineCulture.IetfLanguageTag
    }
    
    $primaryLanguage = checkForLanguage -langId $machinelangId

    $additionalLanguages = @()
    [String[]]$allLanguages = @()

    switch ($Languages) {
      "CurrentOfficeLanguages" 
      {
         if ($officeConfig) {
            $primaryLanguage = $officeConfig.ClientCulture
         } 

         if (!($primaryLanguage)) {
            $msiPrimaryLanguage = msiGetOfficeUILanguage -regProv $regProv
            if ($msiPrimaryLanguage) {
               $primaryLanguage =  $msiPrimaryLanguage
            }

            $primaryLanguage = checkForLanguage -langId $machinelangId
         }
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

         $returnLangs = getLanguages -regProv $regProv

         foreach ($returnLang in $returnLangs) {
            $additionalLanguages += $returnLang
         }
         
      }
    }

    if ($primaryLanguage) {
        $allLanguages += $primaryLanguage.ToLower()
    }

    foreach ($lang in $additionalLanguages) {
      if ($lang.GetType().Name.ToLower().Contains("string")) {
        if ($lang.Contains("-")) {
          [bool]$addLang = $true

          foreach ($language in $allLanguages) {
             if ($language.ToLower() -eq $lang.ToLower()) {
                $addLang = $false
             }
          }

          if ($addLang) {
             $allLanguages += $lang.ToLower()
          }
        }
      }
    }

    if (!($primaryLanguage)) {
            <# write log#>
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Cannot find matching Office language for: $primaryLanguage"
        throw "Cannot find matching Office language for: $primaryLanguage"
    }

    foreach ($productId in $splitProducts) { 
       $excludeApps = $NULL

       if ($Languages -eq "CurrentOfficeLanguages") {
           $additionalLanguages = New-Object System.Collections.ArrayList
       }

       if ($officeConfig.ClickToRunInstalled) {
             $officeKeyPath = $officeConfig.OfficeKeyPath
           
           if ($productId.ToLower().StartsWith("o365")) {
               $excludeApps = odtGetExcludedApps -ConfigDoc $ConfigFile -OfficeKeyPath $officeConfig.OfficeKeyPath -ProductId $productId
           }

           $officeAddLangs = odtGetOfficeLanguages -ConfigDoc $ConfigFile -OfficeKeyPath $officeConfig.OfficeKeyPath -ProductId $productId
       } else {
         if ($officeExists) {
             if($productId.ToLower().StartsWith("o365")) {
                $excludeApps = officeGetExcludedApps -OfficeProducts $officeProducts -computer $computer -Credentials $Credentials
             }
         }
  
         $msiLanguages = msiGetOfficeLanguages -regProv $regProv
         foreach ($msiLanguage in $msiLanguages) {
            $additionalLanguages += $msiLanguage
         }
         
         if (!($additionalLanguages)) {
             foreach ($officeLang in $officeLangs) {
                $additionalLanguages += $officeLang
             }
         }
       }

       if ($officeAddLangs) {
           if (($Languages -eq "CurrentOfficeLanguages") -or ($Languages -eq "AllInUseLanguages")) {
               $additionalLanguages += $officeAddLangs
           }
       }

       if ($additionalLanguages) {
           $additionalLanguages = Get-Unique -InputObject $additionalLanguages -OnType
           
           [bool]$containsLang = $false
           foreach ($additionalLanguage in $additionalLanguages) {
             if ($primaryLanguage) {
               if ($additionalLanguage) {
                  if ($primaryLanguage.ToLower() -eq $additionalLanguage.ToLower()) {
                     $containsLang = $true
                  }
               }
             }
           }
          
           if ($containsLang) {
               $tempLanguages = $additionalLanguages
               $additionalLanguages = New-Object System.Collections.ArrayList
               foreach($tempL in $tempLanguages){
                  if($tempL -ne $primaryLanguage){
                    $additionalLanguages.Add($tempL) | Out-Null
                  }
                  #$additionalLanguages.Remove($primaryLanguage)
               }
           }
       }

       $ChannelName = $NULL
       $ChannelDetect = Detect-Channel
       if ($ChannelDetect) {
          $ChannelName = $ChannelDetect.branch
       }
       
       if ($officeConfig.ClickToRunInstalled) {
          if ($Languages -eq "CurrentOfficeLanguages") {
            
            $officeAddLangs = odtGetOfficeLanguages -ConfigDoc $ConfigFile -OfficeKeyPath $officeConfig.OfficeKeyPath -ProductId $productId
            if ($officeAddLangs) {
               $additionalLanguages = New-Object System.Collections.ArrayList
               $n = 0
               foreach ($language in $officeAddLangs) {
                  if ($n -eq 0) {
                    $primaryLanguage = $language
                  } else {
                    $additionalLanguages += $language
                  }
                  $n++
               }
            }
           
          }
       }

       odtAddProduct -ConfigDoc $ConfigFile -ProductId $productId -ExcludeApps $excludeApps -Version $officeConfig.Version `
                     -Platform $productPlatform -ClientCulture $primaryLanguage -AdditionalLanguages $additionalLanguages -Channel $ChannelName


       if ($officeConfig) {
          if (($officeConfig.UpdatesEnabled) -or ($officeConfig.UpdateUrl) -or  ($officeConfig.UpdateDeadline)) {
            odtAddUpdates -ConfigDoc $ConfigFile -Enabled $officeConfig.UpdatesEnabled -UpdatePath $officeConfig.UpdateUrl -Deadline $officeConfig.UpdateDeadline
          }
       }
    }

    $clickToRunKeys = 'SOFTWARE\Microsoft\Office\ClickToRun',
                        'SOFTWARE\Microsoft\Office\15.0\ClickToRun'

    foreach($key in $clickToRunKeys){
        $configKeys = $regProv.EnumKey($HKLM, $key)
        $clickToRunList = $configKeys.snames
        foreach($list in $clickToRunList){
            if($list -match 'Configuration'){
                $configPath = Join-Path $key "Configuration"
                $sharedLicense = $regProv.GetStringValue($HKLM, $configPath, "SharedComputerLicensing").sValue
                if($sharedLicense -eq '1'){
                   Set-ODTConfigProperties -SharedComputerLicensing "1" -ConfigDoc $ConfigFile
                }
            }
        }
    }  
    
    if ($IncludeUpdatePathAsSourcePath) {
      if ($officeConfig.UpdateUrl) {
          odtSetAdd -ConfigDoc $ConfigFile -SourcePath $officeConfig.UpdateUrl
      }
    }

    $formattedXml = Format-XML ([xml]($ConfigFile)) -indent 4
    # write log
    $lineNum = Get-CurrentLineNumber    
    $filName = Get-CurrentFileName 
    WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Write XML output"
    if (($PSCmdlet.MyInvocation.PipelineLength -eq 1) -or `
        ($PSCmdlet.MyInvocation.PipelineLength -eq $PSCmdlet.MyInvocation.PipelinePosition)) {

        $results = new-object PSObject[] 0;
        $Result = New-Object -TypeName PSObject 
        Add-Member -InputObject $Result -MemberType NoteProperty -Name "ConfigurationXML" -Value $formattedXml

        if ($ComputerName.Length -gt 1) {
            Add-Member -InputObject $Result -MemberType NoteProperty -Name "LanguageIds" -Value $allLanguages
            Add-Member -InputObject $Result -MemberType NoteProperty -Name "ComputerName" -Value $computer
        }

        if ($TargetFilePath) {
           $formattedXml | Out-File -FilePath $TargetFilePath
           if ($ComputerName.Length -eq 1) {
               $Result = $formattedXml
           }
        
        }
        $Result

    } else {
        if ($TargetFilePath) {
           $formattedXml | Out-File -FilePath $TargetFilePath
        }

        $allLanguages = Get-Unique -InputObject $allLanguages

        $results = new-object PSObject[] 0;
        $Result = New-Object -TypeName PSObject 
        Add-Member -InputObject $Result -MemberType NoteProperty -Name "TargetFilePath" -Value $TargetFilePath
        Add-Member -InputObject $Result -MemberType NoteProperty -Name "LanguageIds" -Value $allLanguages
        Add-Member -InputObject $Result -MemberType NoteProperty -Name "ConfigurationXML" -Value $formattedXml
        $Result
    }
    
    #return $ConfigFile
  } catch {
    $errorMessage = $computer + ": " + $_
    Write-Host $errorMessage
    $fileName = $_.InvocationInfo.ScriptName.Substring($_.InvocationInfo.ScriptName.LastIndexOf("\")+1)
    WriteToLogFile -LNumber $_.InvocationInfo.ScriptLineNumber -FName $fileName -ActionError $_
    throw;
  }

  }
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
Version: 1.0.5
DateCreated: 2015-07-01
DateUpdated: 2016-10-14
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

    $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet',[string[]]$defaultDisplaySet)
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

    [System.Collections.ArrayList]$VersionList = New-Object -TypeName System.Collections.ArrayList
    [System.Collections.ArrayList]$PathList = New-Object -TypeName System.Collections.ArrayList
    [System.Collections.ArrayList]$PackageList = New-Object -TypeName System.Collections.ArrayList
    [System.Collections.ArrayList]$ClickToRunPathList = New-Object -TypeName System.Collections.ArrayList
    [System.Collections.ArrayList]$ConfigItemList = New-Object -TypeName  System.Collections.ArrayList
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
            if ($configItems) {
               foreach ($configId in $configItems.sNames) {
                 if ($configId) {
                    $Add = $ConfigItemList.Add($configId.ToUpper())
                 }
               }
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

            [string]$officeLangResourcePath = join-path  $path "Common\LanguageResources"
            $mainLangId = $regProv.GetDWORDValue($HKLM, $officeLangResourcePath, "SKULanguage").uValue
            if ($mainLangId) {
                $mainlangCulture = [globalization.cultureinfo]::GetCultures("allCultures") | where {$_.LCID -eq $mainLangId}
                if ($mainlangCulture) {
                    $cltr.ClientCulture = $mainlangCulture.Name
                }
            }

            [string]$officeLangPath = join-path  $path "Common\LanguageResources\InstalledUIs"
            $langValues = $regProv.EnumValues($HKLM, $officeLangPath);
            if ($langValues) {
               foreach ($langValue in $langValues) {
                  $langCulture = [globalization.cultureinfo]::GetCultures("allCultures") | where {$_.LCID -eq $langValue}
               } 
            }

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
                if ($packageName) {
                   $AddItem = $PackageList.Add($packageName.Replace(' ', '').ToLower())
                }
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
           if (!($installPath)) { continue }
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

           if ($ConfigItemList.Contains($key.ToUpper()) -and $name.ToUpper().Contains("MICROSOFT OFFICE") -and $name.ToUpper() -notlike "*MUI*" -and $name.ToUpper() -notlike "*VISIO*" -and $name.ToUpper() -notlike "*PROJECT*") {
              $primaryOfficeProduct = $true
           }

           $clickToRunComponent = $regProv.GetDWORDValue($HKLM, $path, "ClickToRunComponent").uValue
           $uninstallString = $regProv.GetStringValue($HKLM, $path, "UninstallString").sValue
           if (!($clickToRunComponent)) {
              if ($uninstallString) {
                 if ($uninstallString.Contains("OfficeClickToRun")) {
                     $clickToRunComponent = $true
                 }
              }
           }

           $modifyPath = $regProv.GetStringValue($HKLM, $path, "ModifyPath").sValue 
           $version = $regProv.GetStringValue($HKLM, $path, "DisplayVersion").sValue

           $cltrUpdatedEnabled = $NULL
           $cltrUpdateUrl = $NULL
           $clientCulture = $NULL;

           [string]$clickToRun = $false

           if ($clickToRunComponent) {
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

  $results = Get-Unique -InputObject $results 

  return $results;
}

}

function getCTRConfig() {
    param(
       [Parameter(ValueFromPipelineByPropertyName=$true)]
       $regProv = $NULL
    )

    $HKLM = [UInt32] "0x80000002"
    $computerName = $env:COMPUTERNAME

    if (!($regProv)) {
        $regProv = Get-Wmiobject -list "StdRegProv" -namespace root\default -computername $computerName -ErrorAction Stop
    }
    
    $officeCTRKeys = 'SOFTWARE\Microsoft\Office\15.0\ClickToRun',
                     'SOFTWARE\Wow6432Node\Microsoft\Office\15.0\ClickToRun',
                     'SOFTWARE\Microsoft\Office\ClickToRun',
                     'SOFTWARE\Wow6432Node\Microsoft\Office\ClickToRun'

    $Object = New-Object PSObject
    $Object | add-member Noteproperty ClickToRunInstalled $false

    [string]$officeKeyPath = "";
    foreach ($regPath in $officeCTRKeys) {
       [string]$installPath = $regProv.GetStringValue($HKLM, $regPath, "InstallPath").sValue
       if ($installPath) {
          if ($installPath.Length -gt 0) {
              $officeKeyPath = $regPath;
              break;
          }
       }
    }

    if ($officeKeyPath.Length -gt 0) {
        $Object.ClickToRunInstalled = $true

        $configurationPath = join-path $officeKeyPath "Configuration"

        [string]$platform = $regProv.GetStringValue($HKLM, $configurationPath, "Platform").sValue
        [string]$clientCulture = $regProv.GetStringValue($HKLM, $configurationPath, "ClientCulture").sValue
        [string]$productIds = $regProv.GetStringValue($HKLM, $configurationPath, "ProductReleaseIds").sValue
        [string]$versionToReport = $regProv.GetStringValue($HKLM, $configurationPath, "VersionToReport").sValue
        [string]$updatesEnabled = $regProv.GetStringValue($HKLM, $configurationPath, "UpdatesEnabled").sValue
        [string]$updateUrl = $regProv.GetStringValue($HKLM, $configurationPath, "UpdateUrl").sValue
        [string]$updateDeadline = $regProv.GetStringValue($HKLM, $configurationPath, "UpdateDeadline").sValue

        if (!($productIds)) {
            $productIds = ""
            $officeActivePath = Join-Path $officeKeyPath "ProductReleaseIDs\Active"
            $officeProducts = $regProv.EnumKey($HKLM, $officeActivePath)

            foreach ($productName in $officeProducts.sNames) {
               if ($productName.ToLower() -eq "stream") { continue }
               if ($productName.ToLower() -eq "culture") { continue }
               if ($productIds.Length -gt 0) { $productIds += "," }
               $productIds += "$productName"
            }
        }

        $splitProducts = $productIds.Split(',');

        if ($platform.ToLower() -eq "x86") {
            $platform = "32"
        } else {
            $platform = "64"
        }

        $Object | add-member Noteproperty Platform $platform
        $Object | add-member Noteproperty ClientCulture $clientCulture
        $Object | add-member Noteproperty ProductReleaseIds $productIds
        $Object | add-member Noteproperty Version $versionToReport
        $Object | add-member Noteproperty UpdatesEnabled $updatesEnabled
        $Object | add-member Noteproperty UpdateUrl $updateUrl
        $Object | add-member Noteproperty UpdateDeadline $updateDeadline
        $Object | add-member Noteproperty OfficeKeyPath $officeKeyPath
        
    } 

    return $Object 

}

function getOfficeConfig() {
    param(
       [Parameter(ValueFromPipelineByPropertyName=$true)]
       $regProv = $NULL,
       [Parameter(ValueFromPipelineByPropertyName=$true)]
       [string[]]$ComputerName = $env:COMPUTERNAME,
       [Parameter(ValueFromPipelineByPropertyName=$true)]
       [PSObject]$mainOfficeProduct = $NULL,
       [Parameter(ValueFromPipelineByPropertyName=$true)]
       [PSObject[]]$officeProducts = $NULL
    )

    #HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Office\14.0\Common\InstallRoot
    
    $officeCTRKeys = 'SOFTWARE\Microsoft\Office',
                     'SOFTWARE\Wow6432Node\Microsoft\Office'


    $Object = New-Object PSObject
    $Object | add-member Noteproperty ClickToRunInstalled $false

    [string]$officeKeyPath = "";
    foreach ($regPath in $officeCTRKeys) {
       $officeVersionNums = $regProv.EnumKey($HKLM, $regPath)

       foreach ($officeVersionNum in $officeVersionNums.sNames) {
           [string]$officePath = join-path $regPath "$officeVersionNum\Common\InstallRoot"
           [string]$installPath = $regProv.GetStringValue($HKLM, $officePath, "Path").sValue
           if ($installPath) {
              if ($installPath.Length -gt 0) {
                  $officeKeyPath = join-path $regPath $officeVersionNum
                  break;
              }
           }
       }
    }

    if ($officeKeyPath.Length -gt 0) {
        $Object.ClickToRunInstalled = $false

        $productIds = generateProductReleaseIds -OfficeProducts $officeProducts

        $productDisplayName = ""
        $productBitness = ""
        $productVersion = ""

        if ($officeInstall.Bitness) {
            if ($officeInstall.Bitness.ToLower() -eq "32-bit") {
                $officeInstall.Bitness = "32"
            } else {
                $officeInstall.Bitness = "64"
            }
            $productBitness = $officeInstall.Bitness
            $productDisplayName = $officeInstall.DisplayName
            $productVersion = $officeInstall.Version
        } else {
            if ($mainOfficeProduct) 
            {
               if ($mainOfficeProduct -is [System.Array]) {
                 $firstProduct = $mainOfficeProduct[0]
               } else{
                 $firstProduct = $mainOfficeProduct
               }
               
               if ($firstProduct.Bitness.ToLower() -eq "32-bit") {
                  $firstProduct.Bitness = "32"
               } else {
                  $firstProduct.Bitness = "64"
               }

               $productBitness = $firstProduct.Bitness
               $productDisplayName = $firstProduct.DisplayName
               $productVersion = $firstProduct.Version
            }
        }

        $Object | add-member Noteproperty Platform $productBitness
        $Object | add-member Noteproperty DisplayName $productDisplayName
        $Object | add-member Noteproperty Version $productVersion
        $Object | add-member Noteproperty OfficeKeyPath $officeKeyPath
        $Object | add-member Noteproperty ProductReleaseIds $productIds
    } 

    return $Object 

}

function generateProductReleaseIds() {
    param(
       [Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true, Position=0)]
       [PSObject[]]$MainOfficeProduct = $NULL,

       [Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true, Position=0)]
       [PSObject[]]$OfficeProducts = $NULL
    )

    $productReleaseIds = ""

    if (!($MainOfficeProduct)) 
    {
       $productReleaseIds += "O365ProPlusRetail"
    }

    foreach ($OfficeProduct in $OfficeProducts) 
    {
        if ($OfficeProduct.DisplayName.ToLower().Contains("microsoft") -and
            $OfficeProduct.DisplayName.ToLower().Contains("visio")) {

            if ($productReleaseIds.IndexOf("VisioProRetail") -eq -1) {
                if ($productReleaseIds.Length -gt 0) {
                   $productReleaseIds += ","
                }
                $productReleaseIds += "VisioProRetail"
            }
        }
        if ($OfficeProduct.DisplayName.ToLower().Contains("microsoft") -and
            $OfficeProduct.DisplayName.ToLower().Contains("project")) {

            if ($productReleaseIds.IndexOf("ProjectProRetail") -eq -1) {
                if ($productReleaseIds.Length -gt 0) {
                   $productReleaseIds += ","
                }
                $productReleaseIds += "ProjectProRetail"
            }
        }
        if ($OfficeProduct.DisplayName.ToLower().Contains("microsoft") -and
            $OfficeProduct.DisplayName.ToLower().Contains("sharepoint designer")) {

            if ($productReleaseIds.IndexOf("SPDRetail") -eq -1) {
                if ($productReleaseIds.Length -gt 0) {
                   $productReleaseIds += ","
                }
                $productReleaseIds += "SPDRetail"
            }
        }
    }

    $productReleaseIds = Get-Unique -InputObject $productReleaseIds

    return $productReleaseIds
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
        [System.Collections.ArrayList]$appLanguages1 = New-Object System.Collections.ArrayList

        #SOFTWARE\Wow6432Node\Microsoft\Office\14.0\Common\LanguageResources\InstalledUIs

        $productsPath = join-path $officeKeyPath "ProductReleaseIDs\Active\$ProductId"
        $installedCultures = $regProv.EnumKey($HKLM, $productsPath)
      
        foreach ($installedCulture in $installedCultures.sNames) {
        if($installedCulture){
            if ($installedCulture.Contains("-") -and !($installedCulture.ToLower() -eq "x-none")) {
                $addItem = $appLanguages1.Add($installedCulture) 
            }
            }
        }

        if ($appLanguages1.Count) {
            $productsPath = join-path $officeKeyPath "ProductReleaseIDs\Active\$ProductId"
        } else {
            $productReleasePath = Join-Path $officeKeyPath "ProductReleaseIDs"
            $guids= $regProv.EnumKey($HKLM, $productReleasePath)

            foreach ($guid in $guids.sNames) {

                $productsPath = Join-Path $officeKeyPath "ProductReleaseIDs\$guid\$ProductId.16"
                $installedCultures = $regProv.EnumKey($HKLM, $productsPath)
      
                foreach ($installedCulture in $installedCultures.sNames) {
                   if($installedCulture){
                      if ($installedCulture.Contains("-") -and !($installedCulture.ToLower() -eq "x-none")) {
                            $addItem = $appLanguages1.Add($installedCulture) 
                      }
                   }
                }

                if ($appLanguages1.Count) {
                   $productsPath = Join-Path $officeKeyPath "ProductReleaseIDs\Culture\$ProductId"
                }
 
            }
        }

        return $appLanguages1;
    }
}

function msiGetOfficeUILanguage() {
    param(
       [Parameter(ValueFromPipelineByPropertyName=$true)]
       $regProv = $NULL
    )

    begin {
        $HKU = [UInt32] "0x80000003"
    }

    process {
     $computer = "."

     $msiUILanguages = @()

     $localUsers = $regProv.EnumKey($HKU, "")

     foreach ($localUser in $localUsers.sNames) {
        $regPathLangResource = "$localUser\SOFTWARE\Microsoft\Office\15.0\Common\LanguageResources"
        $UILanguage = $regProv.GetDWordValue($HKU, $regPathLangResource, "UILanguage").uValue

        if ($UILanguage) {
            $langCulture = [globalization.cultureinfo]::GetCultures("allCultures") | where {$_.LCID -eq $UILanguage}
            $convertLang = checkForLanguage -langId $langCulture.Name

            if ($convertLang) {
                $msiUILanguages += $convertLang
            }
        }
     }

     $primaryLanguage = ($msiUILanguages | Group-Object | Sort-Object Count -descending | Select-Object -First 1).Name

     return $primaryLanguage
   }
}

function msiGetOfficeLanguages() {
    param(
       [Parameter(ValueFromPipelineByPropertyName=$true)]
       $regProv = $NULL
    )

    begin {
        $HKU = [UInt32] "0x80000003"
    }

    process {
     $computer = "."

     $msiLanguages = @()

     if (!($regProv)) {
        $regProv = Get-Wmiobject -list "StdRegProv" -namespace root\default -computername $computer -ErrorAction Stop
     }

     $localUsers = $regProv.EnumKey($HKU, "")

     foreach ($localUser in $localUsers.sNames) {
        $regPathLangResource = "$localUser\SOFTWARE\Microsoft\Office\15.0\Common\LanguageResources"
        $regPathEnabledLangs = "$localUser\SOFTWARE\Microsoft\Office\15.0\Common\LanguageResources\EnabledLanguages"

        $UILanguageNum = $regProv.GetDWordValue($HKU, $regPathLangResource, "UILanguage").uValue
        if ($UILanguageNum) {
            $langCulture = [globalization.cultureinfo]::GetCultures("allCultures") | where {$_.LCID -eq $UILanguageNum}
            $UILanguage = checkForLanguage -langId $langCulture
        } else {
            $UILanguage = ""
        }

        $enabledLanguages = $regProv.EnumValues($HKU, $regPathEnabledLangs)

        foreach ($enabledLanguage in $enabledLanguages.sNames) {

           $languageStatus = $regProv.GetStringValue($HKU, $regPathEnabledLangs, $enabledLanguage).sValue
           if($languageStatus){
           if ($languageStatus.ToLower() -eq "on") {
               $langCulture = [globalization.cultureinfo]::GetCultures("allCultures") | where {$_.LCID -eq $enabledLanguage}
               $convertLang = checkForLanguage -langId $langCulture 

               if ($convertLang) {
                   $flgInclude = $true

                   if ($convertLang) {
                       if ($UILanguage) {
                           if ($UILanguage.ToLower() -eq $convertLang.ToLower()) {
                               $flgInclude = $false
                           }
                       }

                       if ($flgInclude) {
                           $msiLanguages += $convertLang
                       }
                   }
               }
           }
           }
        }
     }

     $msiLanguages = $msiLanguages | Get-Unique

     if (!($msiLanguages)) {
        $msiLanguages = @()
     }

     return $msiLanguages

   }
}

function getLanguages() {
    param(
       [Parameter(ValueFromPipelineByPropertyName=$true)]
       $regProv = $NULL
    )

  $returnLangs = @() 

  $HKU = [UInt32] "0x80000003"
  $userKeys = $regProv.EnumKey($HKU, "");

  foreach ($userKey in $userKeys.sNames) {
     if ($userKey.Length -gt 8 -and !($userKey.ToLower().EndsWith("_classes"))) {
       [string]$userProfilePath = join-path $userKey "Control Panel\International\User Profile"
       [string[]]$userLanguages = $regProv.GetMultiStringValue($HKU, $userProfilePath, "Languages").sValue
       foreach ($userLang in $userLanguages) {
       if($userLang){
         $convertLang = checkForLanguage -langId $userLang 
         }
         if ($convertLang) {
             $returnLangs += $convertLang.ToLower()
         }
       }
        
     }
  }
  
  $langPacks = $regProv.EnumKey($HKLM, "SYSTEM\CurrentControlSet\Control\MUI\UILanguages");
  foreach ($langPackName in $langPacks.sNames) {
     [bool]$addReturnLang = $true

     foreach ($returnLang in $returnLangs) {
        if ($returnLang.ToLower() -eq $langPackName.ToLower()) {
           $addReturnLang = $false
        }
     }

     if ($addReturnLang) {
        $returnLangs += $langPackName.ToLower() 
     }
  }

  #HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\MUI\UILanguages

  if ($returnLangs.Count -gt 1) {
     $returnLangs = Get-Unique -InputObject $returnLangs
  }

  $validLangs = @()
  foreach($lang in $returnlangs){
    $langExists = $false
    foreach ($tmpLang in $availablelangs) {
       if ($tmpLang) {
          if ($lang) {
              if ($tmpLang.ToLower() -eq $lang.ToLower()) {
                 $langExists = $true
              }
          }
       }
    }

    if($langExists){
        $validLangs += $lang
    }   
  }
  
  $returnLangs = $validLangs
  return $returnLangs
}

function checkForLanguage() {
    param(
       [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
       [string]$langId = $NULL
    )

    [bool]$langExists = $false
    foreach ($availableLang in $availableLangs) {
       if ($availableLang.ToLower() -eq $langId.Trim().ToLower()) {
          $langExists = $true
       }
    }

    if ($langExists) {
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

function officeGetExcludedApps() {
    param(
       [Parameter(ValueFromPipelineByPropertyName=$true, Position=0)]
       [PSObject[]]$OfficeProducts = $NULL,

       [string]$Credentials,

       [string]$computer = $env:COMPUTERNAME
    )

    begin {
        $HKLM = [UInt32] "0x80000002"
        $HKCR = [UInt32] "0x80000000"

        $allExcludeApps = 'Access','Excel','Groove','InfoPath','OneNote','Outlook',
                       'PowerPoint','Publisher','Word'

        if ($Credentials) {
            $regProv = Get-Wmiobject -list "StdRegProv" -namespace root\default -computername $computer -Credential $Credentials  -ErrorAction Stop
            $os = Get-WMIObject win32_operatingsystem -computername $computer -Credential $Credentials -ErrorAction Stop
        } 
        else {
            $regProv = Get-Wmiobject -list "StdRegProv" -namespace root\default -computername $computer  -ErrorAction Stop
            $os = Get-WMIObject win32_operatingsystem -computername $computer  -ErrorAction Stop
        }
    }

    process{
        $OfficeVersion = Get-OfficeVersion -ComputerName $computer
        $OfficeVersion = $OfficeVersion.Version.Split(".")[0]

        switch($os.OSArchitecture){
            "32-bit"
            {
                $osBitness = '32'
                $appKeyPath = 'SOFTWARE\Microsoft\Office'
            }
            "64-bit"
            {
                $osBitness = '64'
                $appKeyPath = 'SOFTWARE\WOW6432Node\Microsoft\Office'

            }
        }
        
        switch($OfficeVersion){
            "11"
            {
                $bitPath = '11.0'
            }
            "12"
            {
                $bitPath = '12.0'
                
            }
            "14"
            {
                $bitPath = '14.0'
            }
            "15"
            {
                $bitPath = '15.0'
            } 
        }

        $appKeyPath = Join-Path $appKeyPath $bitPath
         
        $appKeys = $regProv.EnumKey($HKLM, $appKeyPath)
        $appList = $appKeys.sNames

        $appsToExclude = @()

        foreach($appName in $allExcludeApps){
            [bool]$appInstalled = $false

            foreach ($OfficeProduct in $appList){
                if($OfficeProduct.ToLower() -like $appName.ToLower()){
                    if($OfficeProduct -eq "OneNote"){
                        $onRegPath = Join-Path $appKeyPath $OfficeProduct
                        $onInstallKey = $regProv.EnumKey($HKLM, $onRegPath)
                        $onRegKeys = $onInstallKey.sNames
                        foreach($key in $onRegKeys){
                            if($key -like "InstallRoot"){
                                $onInstallRegKey = Join-Path $onRegPath "InstallRoot"
                                $installRoot = $regProv.GetStringValue($HKLM, $onInstallRegKey, "Path").sValue
                                $pathChk = Test-Path -Path $installRoot
                                if($pathChk){
                                    $appInstalled = $true
                                    break;
                                }
                            }
                        }              
                    }
                    else{
                        $appInstalled = $true
                        break;
                    }
                }
            }

            if(!($appInstalled)){
                $appsToExclude += $appName
            }
        }
            
        return $appsToExclude;
    }
}

function officeGetLanguages() {
   param(
       [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
       $regProv = $NULL,
       [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
       [string]$OfficeKeyPath = $NULL
   )

   $HKLM = [UInt32] "0x80000002"
   $HKCR = [UInt32] "0x80000000"

   [string]$officeLangPath = join-path $OfficeKeyPath "Common\LanguageResources\InstalledUIs"

   [System.Collections.ArrayList] $returnLangs = New-Object System.Collections.ArrayList

   $langValues = $regProv.EnumValues($HKLM, $officeLangPath);
 
   foreach ($langValue in $langValues.sNames) {
        $langCulture = [globalization.cultureinfo]::GetCultures("allCultures") | where {$_.LCID -eq $langValue}     
        $convertLang = checkForLanguage -langId $langCulture 
        if ($convertLang) {
            $returnLangs.Add($convertLang.ToLower()) | Out-Null
        }
   }
  
   if ($returnLangs.Count -gt 1) {
     $returnLangs = $returnLangs | Get-Unique 
   }

   return $returnLangs

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

        $allExcludeApps = 'Access','Excel','Groove','InfoPath','Lync','OneDrive','OneNote','Outlook',
                       'PowerPoint','Publisher','Word'
        #"SharePointDesigner","Visio", 'Project'
    }

    process {
        $configPath = join-path $officeKeyPath "Configuration"

        $appsToExclude = @() 

        $keyValues = $regProv.EnumValues($HKLM, $configPath)

        foreach ($keyValue in $keyValues.sNames) {
            $checkValue = $ProductId + ".ExcludedApps"

            if ($keyValue.ToLower() -eq $checkValue.ToLower()) {
                $excludeApps = $regProv.GetStringValue($HKLM, $configPath, $checkValue).sValue

                $appSplit = $excludeApps.Split(',');

                foreach ($app in $appSplit){
                    $app = (Get-Culture).textinfo.totitlecase($app)

                    $appsToExclude += $app
                }
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
       [string]$Version = $NULL,

       [Parameter(ValueFromPipelineByPropertyName=$true)]
       [string]$Channel = $NULL
    )

    [System.XML.XMLElement]$ConfigElement=$NULL
    if($ConfigDoc.Configuration -eq $null){
        $ConfigElement=$ConfigDoc.CreateElement("Configuration")
        $ConfigDoc.appendChild($ConfigElement) | Out-Null
    }

    [System.XML.XMLElement]$AddElement=$NULL
    if($ConfigFile.Configuration.Add -eq $null){
        $AddElement=$ConfigDoc.CreateElement("Add")
        $ConfigDoc.DocumentElement.appendChild($AddElement) | Out-Null
    } else {
        $AddElement = $ConfigDoc.Configuration.Add 
    }

    if ($Version) {
       if ($Version.StartsWith("16.")) {
          $AddElement.SetAttribute("Version", $Version) | Out-Null
       }
    }

    if ($Channel) {
       $AddElement.SetAttribute("Channel", $Channel) | Out-Null
    }

    if ($Platform) {
       $AddElement.SetAttribute("OfficeClientEdition", $Platform) | Out-Null
    }

    [System.XML.XMLElement]$ProductElement = $ConfigDoc.Configuration.Add.Product | where { $_.ID -eq $ProductId }
    if($ProductId){
    if($ProductElement -eq $null){
        [System.XML.XMLElement]$ProductElement=$ConfigDoc.CreateElement("Product")
        $AddElement.appendChild($ProductElement) | Out-Null
        $ProductElement.SetAttribute("ID", $ProductId) | Out-Null
    }
    }

    $LanguageIds = @($ClientCulture)

    foreach ($addLang in $AdditionalLanguages) {
       $LanguageIds += $addLang 
    }

    foreach($LanguageId in $LanguageIds){    
       if ($LanguageId) {
          if ($LanguageId.Length -gt 0) {
            [System.XML.XMLElement]$LanguageElement = $ProductElement.Language | where { $_.ID -eq $LanguageId }
            if($LanguageElement -eq $null){
                [System.XML.XMLElement]$LanguageElement=$ConfigFile.CreateElement("Language")
                $ProductElement.appendChild($LanguageElement) | Out-Null
                $LanguageElement.SetAttribute("ID", $LanguageId.ToString().ToLower()) | Out-Null
            }
          }
       }
    }

    foreach($ExcludeApp in $ExcludeApps){
    if($ExcludeApp){
        [System.XML.XMLElement]$ExcludeAppElement = $ProductElement.ExcludeApp | where { $_.ID -eq $ExcludeApp }
        if($ExcludeAppElement -eq $null){
            [System.XML.XMLElement]$ExcludeAppElement=$ConfigDoc.CreateElement("ExcludeApp")
            $ProductElement.appendChild($ExcludeAppElement) | Out-Null
            $ExcludeAppElement.SetAttribute("ID", $ExcludeApp) | Out-Null
        }
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
        <# write log#>
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "no configuration element"
            throw $NoConfigurationElement
        }
        [bool]$addUpdates = $false
        $hasEnabled = $false
        if($Enabled){$hasEnabled = $true}else{$hasEnabled = $false}
        
        $hasUpdatePath = $false
        if($UpdatePath){$hasUpdatePath = $true}else{$hasUpdatePath = $false}
        if(($hasEnabled -eq $true) -or ($hasUpdatePath -eq $true)){
           $addUpdates = $true
        }

        if ($addUpdates) {
            #Get the Updates Element if it exists
            [System.XML.XMLElement]$UpdateElement = $ConfigDoc.Configuration.GetElementsByTagName("Updates").Item(0)
            if($ConfigDoc.Configuration.Updates -eq $null){
                [System.XML.XMLElement]$UpdateElement=$ConfigDoc.CreateElement("Updates")
                $ConfigDoc.Configuration.appendChild($UpdateElement) | Out-Null
            }

            #Set the desired values
            if($Enabled){
                $UpdateElement.SetAttribute("Enabled", $Enabled.ToString().ToUpper()) | Out-Null
            } else {
              if ($PSBoundParameters.ContainsKey('Enabled')) {
                 if ($ConfigDoc.Configuration.Updates) {
                     $ConfigDoc.Configuration.Updates.RemoveAttribute("Enabled")
                 }
              }
            }

            if($UpdatePath){
                $UpdateElement.SetAttribute("UpdatePath", $UpdatePath) | Out-Null
            } else {
              if ($PSBoundParameters.ContainsKey('UpdatePath')) {
                 if ($ConfigDoc.Configuration.Updates) {
                     $ConfigDoc.Configuration.Updates.RemoveAttribute("UpdatePath")
                 }
              }
            }

            if($TargetVersion){
                $UpdateElement.SetAttribute("TargetVersion", $TargetVersion) | Out-Null
            } else {
              if ($PSBoundParameters.ContainsKey('TargetVersion')) {
                 if ($ConfigDoc.Configuration.Updates) {
                     $ConfigDoc.Configuration.Updates.RemoveAttribute("TargetVersion")
                 }
              }
            }

            if($Deadline){
                $UpdateElement.SetAttribute("Deadline", $Deadline) | Out-Null
            } else {
              if ($PSBoundParameters.ContainsKey('Deadline')) {
                 if ($ConfigDoc.Configuration.Updates) {
                     $ConfigDoc.Configuration.Updates.RemoveAttribute("Deadline")
                 }
              }
            }
        }
       

    }
}

Function odtSetAdd{

    Param(

        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [System.XML.XMLDocument]$ConfigDoc = $NULL,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $SourcePath = $NULL,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $Version,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $Bitness

    )

    Process{
        #Check for proper root element
        if($ConfigDoc.Configuration -eq $null){
        <# write log#>
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "no configuration element"
            throw $NoConfigurationElement
        }

        #Get Add element if it exists
        if($ConfigDoc.Configuration.Add -eq $null){
            [System.XML.XMLElement]$AddElement=$ConfigFile.CreateElement("Add")
            $ConfigDoc.Configuration.appendChild($AddElement) | Out-Null
        }

        #Set values as desired
        if($SourcePath){
            $ConfigFile.Configuration.Add.SetAttribute("SourcePath", $SourcePath) | Out-Null
        } else {
            if ($PSBoundParameters.ContainsKey('SourcePath')) {
                $ConfigDoc.Configuration.Add.RemoveAttribute("SourcePath")
            }
        }

        if($Version){
            $ConfigDoc.Configuration.Add.SetAttribute("Version", $Version) | Out-Null
        } else {
            if ($PSBoundParameters.ContainsKey('Version')) {
                $ConfigDoc.Configuration.Add.RemoveAttribute("Version")
            }
        }

        if($Bitness){
            $ConfigDoc.Configuration.Add.SetAttribute("OfficeClientEdition", $Bitness) | Out-Null
        } else {
            if ($PSBoundParameters.ContainsKey('OfficeClientEdition')) {
                $ConfigDoc.Configuration.Add.RemoveAttribute("OfficeClientEdition")
            }
        }
    }

}

Function Set-ODTConfigProperties{
<#
.SYNOPSIS
Modifies an existing configuration xml file to set property values
.PARAMETER AutoActivate
If AUTOACTIVATE is set to 1, the specified products will attempt to activate automatically. 
If AUTOACTIVATE is not set, the user may see the Activation Wizard UI.
You must not set AUTOACTIVATE for Office 365 Click-to-Run products. 
.PARAMETER ForceAppShutDown
An installation or removal action may be blocked if Office applications are running. 
Normally, such cases would start a process killer UI. Administrators can set 
FORCEAPPSHUTDOWN value to TRUE to prevent dependence on user interaction. When 
FORCEAPPSHUTDOWN is set to TRUE, any applications that block the action will be shut 
down. Data loss may occur. When FORCEAPPSHUTDOWN is set to FALSE (default), the 
action may fail if Office applications are running.
.PARAMETER PackageGUID
Optional. By default, all Office 2013 App-V packages created by using the Office 
Deployment Tool share the same App-V Package ID. Administrators can use PACKAGEGUID 
to specify a different Package ID. Also, PACKAGEGUID needs to be at least 25 
characters in length and be separated into 5 sections, with each section separated by 
a dash. The sections need to have the following number of characters: 8, 4, 4, 4, and 12. 
.PARAMETER SharedComputerLicensing
Optional. Set SharedComputerLicensing to 1 if you deploy Office 365 ProPlus to shared 
computers by using Remote Desktop Services.
.PARAMETER TargetFilePath
Full file path for the file to be modified and be output to.
.Example
Set-ODTConfigProperties -AutoActivate "1" -TargetFilePath "$env:Public/Documents/config.xml"
Sets config to automatically activate the products
.Example
Set-ODTConfigProperties -ForceAppShutDown "True" -PackageGUID "12345678-ABCD-1234-ABCD-1234567890AB" -TargetFilePath "$env:Public/Documents/config.xml"
Sets the config so that apps are forced to shutdown during install and the package guid
to "12345678-ABCD-1234-ABCD-1234567890AB"
.Notes
Here is what the portion of configuration file looks like when modified by this function:
<Configuration>
  ...
  <Property Name="AUTOACTIVATE" Value="1" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
  <Property Name="PACKAGEGUID" Value="12345678-ABCD-1234-ABCD-1234567890AB" />
  <Property Name="SharedComputerLicensing" Value="0" />
  ...
</Configuration>
#>
    [CmdletBinding()]
    Param(

       [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
       [System.XML.XMLDocument]$ConfigDoc = $NULL,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $AutoActivate,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $ForceAppShutDown,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $PackageGUID,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $SharedComputerLicensing,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $TargetFilePath
    )

    Process{
        #Load file
        [System.XML.XMLDocument]$ConfigFile = $ConfigDoc

        #Check for proper root element
        if($ConfigFile.Configuration -eq $null){
        <# write log#>
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "no configuration element"
            throw $NoConfigurationElement
        }

        #Set each property as desired
        if(($AutoActivate)){
            [System.XML.XMLElement]$AutoActivateElement = $ConfigFile.Configuration.Property | where { $_.Name -eq "AUTOACTIVATE" }
            if($AutoActivateElement -eq $null){
                [System.XML.XMLElement]$AutoActivateElement=$ConfigFile.CreateElement("Property")
            }
                
            $ConfigFile.Configuration.appendChild($AutoActivateElement) | Out-Null
            $AutoActivateElement.SetAttribute("Name", "AUTOACTIVATE") | Out-Null
            $AutoActivateElement.SetAttribute("Value", $AutoActivate) | Out-Null
        }

        if(($ForceAppShutDown)){
            [System.XML.XMLElement]$ForceAppShutDownElement = $ConfigFile.Configuration.Property | where { $_.Name -eq "FORCEAPPSHUTDOWN" }
            if($ForceAppShutDownElement -eq $null){
                [System.XML.XMLElement]$ForceAppShutDownElement=$ConfigFile.CreateElement("Property")
            }
                
            $ConfigFile.Configuration.appendChild($ForceAppShutDownElement) | Out-Null
            $ForceAppShutDownElement.SetAttribute("Name", "FORCEAPPSHUTDOWN") | Out-Null
            $ForceAppShutDownElement.SetAttribute("Value", $ForceAppShutDown) | Out-Null
        }

        if(($PackageGUID)){
            [System.XML.XMLElement]$PackageGUIDElement = $ConfigFile.Configuration.Property | where { $_.Name -eq "PACKAGEGUID" }
            if($PackageGUIDElement -eq $null){
                [System.XML.XMLElement]$PackageGUIDElement=$ConfigFile.CreateElement("Property")
            }
                
            $ConfigFile.Configuration.appendChild($PackageGUIDElement) | Out-Null
            $PackageGUIDElement.SetAttribute("Name", "PACKAGEGUID") | Out-Null
            $PackageGUIDElement.SetAttribute("Value", $PackageGUID) | Out-Null
        }

        if(($SharedComputerLicensing)){
            [System.XML.XMLElement]$SharedComputerLicensingElement = $ConfigFile.Configuration.Property | where { $_.Name -eq "SharedComputerLicensing" }
            if($SharedComputerLicensingElement -eq $null){
                [System.XML.XMLElement]$SharedComputerLicensingElement=$ConfigFile.CreateElement("Property")
            }
                
            $ConfigFile.Configuration.appendChild($SharedComputerLicensingElement) | Out-Null
            $SharedComputerLicensingElement.SetAttribute("Name", "SharedComputerLicensing") | Out-Null
            $SharedComputerLicensingElement.SetAttribute("Value", $SharedComputerLicensing) | Out-Null
        }

    }
} 

Function GetFilePath() {
    Param(
       [Parameter(ValueFromPipelineByPropertyName=$true)]
       [string] $TargetFilePath
    )

    if (!($TargetFilePath)) {
        $TargetFilePath = $global:saveLastFilePath
    }  

    if (!($TargetFilePath)) {
       Write-Host "Enter the path to the XML Configuration File: " -NoNewline
       $TargetFilePath = Read-Host
    } else {
       #Write-Host "Target XML Configuration File: $TargetFilePath"
    }

    return $TargetFilePath
}

Function GetScriptPath() {
 process {
     [string]$scriptPath = "."

     if ($PSScriptRoot) {
       $scriptPath = $PSScriptRoot
     } else {
       $scriptPath = (Get-Item -Path ".\").FullName
     }

     return $scriptPath
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

function Win7Join([string]$st1, [string]$st2){
    [string]$tempStr = $st1 + "\" + $st2
    return $tempStr
}


function Detect-Channel {
   param( 

   )

Process {      
   $channelXml = Get-ChannelXml

   $UpdateChannel = (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration -Name UpdateChannel -ErrorAction SilentlyContinue).UpdateChannel      
   $GPOUpdatePath = (Get-ItemProperty HKLM:\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate -Name updatepath -ErrorAction SilentlyContinue).updatepath
   $GPOUpdateBranch = (Get-ItemProperty HKLM:\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate -Name UpdateBranch -ErrorAction SilentlyContinue).UpdateBranch
   $GPOUpdateChannel = (Get-ItemProperty HKLM:\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate -Name UpdateChannel -ErrorAction SilentlyContinue).UpdateChannel      
   $UpdateUrl = (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration -Name UpdateUrl -ErrorAction SilentlyContinue).UpdateUrl
   $currentBaseUrl = Get-OfficeCDNUrl

   $CurrentChannel = $channelXml.UpdateFiles.baseURL | Where {$_.URL -eq $currentBaseUrl -and $_.branch -notmatch 'Business' }
      
   if($UpdateUrl -ne $null -and $UpdateUrl -like '*officecdn.microsoft.com*'){
       $CurrentChannel = $channelXml.UpdateFiles.baseURL | Where {$_.URL -eq $UpdateUrl -and $_.branch -notmatch 'Business' }  
   }

   if($GPOUpdateChannel -ne $null){
     $CurrentChannel = $channelXml.UpdateFiles.baseURL | ? {$_.branch.ToLower() -eq $GPOUpdateChannel.ToLower()}         
   }

   if($GPOUpdateBranch -ne $null){
     $CurrentChannel = $channelXml.UpdateFiles.baseURL | ? {$_.branch.ToLower() -eq $GPOUpdateBranch.ToLower()}  
   }

   if($GPOUpdatePath -ne $null -and $GPOUpdatePath -like '*officecdn.microsoft.com*'){
     $CurrentChannel = $channelXml.UpdateFiles.baseURL | Where {$_.URL -eq $GPOUpdatePath -and $_.branch -notmatch 'Business' }  
   }

   if($UpdateChannel -ne $null -and $UpdateChannel -like '*officecdn.microsoft.com*'){
     $CurrentChannel = $channelXml.UpdateFiles.baseURL | Where {$_.URL -eq $UpdateChannel -and $_.branch -notmatch 'Business' }  
   }

   return $CurrentChannel
}

}

function Get-ChannelUrl() {
   [CmdletBinding()]
   param( 
      [Parameter(Mandatory=$true)]
      [Channel]$Channel
   )

   Process {
      $channelXml = Get-ChannelXml

      $currentChannel = $channelXml.UpdateFiles.baseURL | Where {$_.branch -eq $Channel.ToString() }
      return $currentChannel
   }
}

function Get-ChannelXml() {
   [CmdletBinding()]
   param( 
    [Parameter()]
    [string]$LogFilePath = "$env:temp\RollBackLogFile.log"  
   )

   process {
       $XMLFilePath = "$PSScriptRoot\ofl.cab"
       Write-Logfile "Line 520: XMLFilePath set to $XMLFilePath"

       if (!(Test-Path -Path $XMLFilePath)) {
           $webclient = New-Object System.Net.WebClient
           $XMLFilePath = "$env:TEMP/ofl.cab"
           $XMLDownloadURL = "http://officecdn.microsoft.com/pr/wsus/ofl.cab"
           $webclient.DownloadFile($XMLDownloadURL,$XMLFilePath)
       }

       if($PSVersionTable.PSVersion.Major -ge '3'){
           $tmpName = "o365client_64bit.xml"
           expand $XMLFilePath $env:TEMP -f:$tmpName | Out-Null
           $tmpName = $env:TEMP + "\o365client_64bit.xml"
       }else {
           $scriptPath = GetScriptPath
           $tmpName = $scriptPath + "\o365client_64bit.xml"           
       }

       [xml]$channelXml = Get-Content $tmpName

       return $channelXml
   }

}

Function Set-OfficeCDNUrl() {
   [CmdletBinding()]
   param( 
      [Parameter(Mandatory=$true)]
      [Channel]$Channel
   )

   Process {
        $CDNBaseUrl = (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration -Name CDNBaseUrl -ErrorAction SilentlyContinue).CDNBaseUrl
        if (!($CDNBaseUrl)) {
           $CDNBaseUrl = (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration -Name CDNBaseUrl -ErrorAction SilentlyContinue).CDNBaseUrl
        }

        $path15 = 'HKLM:\SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration'
        $path16 = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
        $regPath = $path16

        if (Test-Path -Path $path16) { $regPath = $path16 }
        if (Test-Path -Path $path15) { $regPath = $path15 }

        $ChannelUrl = Get-ChannelUrl -Channel $Channel
           
        New-ItemProperty $regPath -Name CDNBaseUrl -PropertyType String -Value $ChannelUrl.URL -Force | Out-Null
   }
}

Function Get-OfficeCDNUrl() {
    $CDNBaseUrl = (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration -Name CDNBaseUrl -ErrorAction SilentlyContinue).CDNBaseUrl
    if (!($CDNBaseUrl)) {
       $CDNBaseUrl = (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration -Name CDNBaseUrl -ErrorAction SilentlyContinue).CDNBaseUrl
    }
    if (!($CDNBaseUrl)) {
        Push-Location
        $path15 = 'HKLM:\SOFTWARE\Microsoft\Office\15.0\ClickToRun\ProductReleaseIDs\Active\stream'
        $path16 = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\ProductReleaseIDs\Active\stream'
        if (Test-Path -Path $path16) { Set-Location $path16 }
        if (Test-Path -Path $path15) { Set-Location $path15 }
        
        try {
        $items = Get-Item . | Select-Object -ExpandProperty property -ErrorAction SilentlyContinue
        if ($items) {
            $properties = $items | ForEach-Object {
               New-Object psobject -Property @{"property"=$_; "Value" = (Get-ItemProperty -Path . -Name $_).$_}
            }

            $value = $properties | Select Value
            $firstItem = $value[0]
            [string] $cdnPath = $firstItem.Value

            $CDNBaseUrl = Select-String -InputObject $cdnPath -Pattern "http://officecdn.microsoft.com/.*/.{8}-.{4}-.{4}-.{4}-.{12}" -AllMatches | % { $_.Matches } | % { $_.Value }
        }
        } catch { }
        Pop-Location
    }
    return $CDNBaseUrl
}

Function WriteToLogFile() {
    param( 
      [Parameter(Mandatory=$true)]
      [string]$LNumber,
      [Parameter(Mandatory=$true)]
      [string]$FName,
      [Parameter(Mandatory=$true)]
      [string]$ActionError
    )
    try{
        $headerString = "Time".PadRight(30, ' ') + "Line Number".PadRight(15,' ') + "FileName".PadRight(60,' ') + "Action"
        $stringToWrite = $(Get-Date -Format G).PadRight(30, ' ') + $($LNumber).PadRight(15, ' ') + $($FName).PadRight(60,' ') + $ActionError

        #check if file exists, create if it doesn't
        $getCurrentDatePath = "C:\Windows\Temp\" + (Get-Date -Format u).Substring(0,10)+"OfficeAutoScriptLog.txt"
        if(Test-Path $getCurrentDatePath){#if exists, append  
             Add-Content $getCurrentDatePath $stringToWrite
        }
        else{#if not exists, create new
             Add-Content $getCurrentDatePath $headerString
             Add-Content $getCurrentDatePath $stringToWrite
        }
    } catch [Exception]{
        Write-Host $_
    }
}

function Get-CurrentLineNumber {
    $MyInvocation.ScriptLineNumber
}

function Get-CurrentFileName{
    $MyInvocation.ScriptName.Substring($MyInvocation.ScriptName.LastIndexOf("\")+1)
}

function Get-CurrentFunctionName {
    (Get-Variable MyInvocation -Scope 1).Value.MyCommand.Name;
}


$availableLangs = @("en-us",
"ar-sa","bg-bg","zh-cn","zh-tw","hr-hr","cs-cz","da-dk","nl-nl","et-ee",
"fi-fi","fr-fr","de-de","el-gr","he-il","hi-in","hu-hu","id-id","it-it",
"ja-jp","kk-kz","ko-kr","lv-lv","lt-lt","ms-my","nb-no","pl-pl","pt-br",
"pt-pt","ro-ro","ru-ru","sr-latn-rs","sk-sk","sl-si","es-es","sv-se","th-th",
"tr-tr","uk-ua","vi-vn",#end of core languages
"af-za","sq-al","am-et","hy-am","as-in","az-latn-az","eu-es","be-by","bn-bd","bn-in","bs-latn-ba","ca-es","prs-af","fil-ph","gl-es","ka-ge","gu-in","is-is","ga-ie","kn-in", #beginning of partial languages
"km-kh","sw-ke","kok-in","ky-kg","lb-lu","mk-mk","ml-in","mt-mt","mi-nz","mr-in","mn-mn","ne-np","nn-no","or-in","fa-ir","pa-in","quz-pe","gd-gb","sr-cyrl-rs","sr-cyrl-ba",#end of partial langauges
"ha-latn-ng","ig-ng","xh-za","zu-za","rw-rw","ps-af","rm-ch","nso-za","tn-za","wo-sn","yo-ng");#proofing languages
