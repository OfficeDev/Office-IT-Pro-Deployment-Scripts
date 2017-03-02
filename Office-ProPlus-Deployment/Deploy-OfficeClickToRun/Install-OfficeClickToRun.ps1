try {
Add-Type  -ErrorAction SilentlyContinue -TypeDefinition @"
   public enum OfficeCTRVersion
   {
      Office2013,
      Office2016
   }
"@
} catch {}

try {
$enum = "
using System;
 
namespace Microsoft.Office
{
     [FlagsAttribute]
     public enum Products
     {
         Unknown = 0,
         O365ProPlusRetail = 1,
         O365BusinessRetail = 2,
         VisioProRetail = 4,
         ProjectProRetail = 8,
         SPDRetail = 16,
         VisioProXVolume = 32,
         VisioStdXVolume = 64,
         ProjectProXVolume = 128,
         ProjectStdXVolume = 256,
         InfoPathRetail = 512,
         SkypeforBusinessEntryRetail = 1024,
         LyncEntryRetail = 2048,
     }
}
"
Add-Type -TypeDefinition $enum -ErrorAction SilentlyContinue
} catch {}

try {
$enum2 = "
using System;
 
    [FlagsAttribute]
    public enum LogLevel
    {
        None=0,
        Full=1
    }
"
Add-Type -TypeDefinition $enum2 -ErrorAction SilentlyContinue
} catch {}

try {
Add-Type  -ErrorAction SilentlyContinue -TypeDefinition @"
   public enum PinAction
   {
      PinToStartMenu,
      PinToTaskbar,
      UnpinFromStartMenu,
      UnpinFromTaskbar
   }
"@
} catch {}

function Install-OfficeClickToRun {
<#
.SYNOPSIS
Installs Office Click-To-Run

.DESCRIPTION
Installs Office Click-To-Run using a specified configuration file or targetfilepath.

.PARAMETER ConfigurationXML
The path to a pre-configured configuration.xml file used for installation.

.PARAMETER TargetFilePath
If no ConfigurationXML is specified, this is the path where the generated configuration.xml will be saved.

.PARAMETER PinToStart 
If $true, all Office apps will be pinned to the Start Menu in Windows 10.

.PARAMETER OfficeVersion
The version of Office Click-To-Run to install. Available options are Office2013 and Office2016. 

.PARAMETER WaitForInstallToFinish
If $true, the PowerShell console will remain open and provide status updates until Office is finished installing.

.PARAMETER PinToStartMenu
Choose one or multiple Office applications to pin to the Start Menu after the installation is finished. 

.PARAMETER PinToTaskbar
Choose one or multiple Office applications to pin to the Taskbar after the installation is finished. Pinning applications
to the Taskbar in Windows 10 is not natively supported.

.EXAMPLE
Install-OfficeClickToRun -ConfigurationXML C:\OfficeDeployment\configuration.xml
Office 2016 Click-To-Run will be installed using the settings in the specified configuration.xml.

.EXAMPLE
Install-OfficeClickToRun -TargetFilePath $env:temp\configuration.xml
Office 2016 Click-To-Run will be installed using an auto-generated configuration file that will be saved to the temp directory.

.EXAMPLE
Install-OfficeClickToRun -TargetFilePath $env:temp\configuration.xml -PinToStartMenu Word,Excel,Outlook -WaitForInstallToFinish $false
Office 2016 Click-To-Run will be installed using an auto-generated configuration file that will be saved to the temp directory. Microsoft
Word, Excel, and Outlook will be pinned to the Start Menu. The PowerShell console will not provide status updates of the installation.
#>
    [CmdletBinding()]
    Param(
        [Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true, Position=0)]
        [string] $ConfigurationXML = $NULL,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $TargetFilePath = $NULL,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [OfficeCTRVersion] $OfficeVersion = "Office2016",

        [Parameter()]
        [bool] $WaitForInstallToFinish = $true,

        [Parameter()]
        [ValidateSet("AllOfficeApps","None","Word","Excel","PowerPoint","OneNote","Access","Publisher","Outlook","Skype for Business",
                     "OneDrive for Business","Project","Visio")]
        [string[]]$PinToStartMenu,

        [Parameter()]
        [ValidateSet("AllOfficeApps","None","Word","Excel","PowerPoint","OneNote","Access","Publisher","Outlook","Skype for Business",
                     "OneDrive for Business","Project","Visio")]
        [string[]]$PinToTaskbar,

        [Parameter()]
        [bool]$InstallProofingTools = $false

    )

    $scriptRoot = GetScriptRoot
	#write log
    $lineNum = Get-CurrentLineNumber    
    $filName = Get-CurrentFileName 
    WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "install office function, loading config file"

    #Load the file
    [System.XML.XMLDocument]$ConfigFile = New-Object System.XML.XMLDocument
        
    if ($TargetFilePath) {
        $ConfigFile.Load($TargetFilePath) | Out-Null
    } else {
        if ($ConfigurationXml) 
        {
            $ConfigFile.LoadXml($ConfigurationXml) | Out-Null
            $global:saveLastConfigFile = $NULL
            $TargetFilePath = $NULL
        }
    }

    [string]$officeCtrPath = ""

    if ($OfficeVersion -eq "Office2013") {
        $officeCtrPath = Join-Path $scriptRoot "Office2013Setup.exe"
        if (!(Test-Path -Path $officeCtrPath)) {
           throw "Cannot find the Office 2013 Setup executable"
        }
    }

    if ($OfficeVersion -eq "Office2016") {
        $officeCtrPath = $scriptRoot + "\Office2016Setup.exe"
        if (!(Test-Path -Path $officeCtrPath)) {
           throw "Cannot find the Office 2016 Setup executable"
        }
    }
    
    if (!($TargetFilePath)) {
      if ($ConfigurationXML) {
         $TargetFilePath = $scriptRoot + "\configuration.xml"
         New-Item -Path $TargetFilePath -ItemType "File" -Value $ConfigurationXML -Force | Out-Null
      }
    }
    
    if (!(Test-Path -Path $TargetFilePath)) {
       $TargetFilePath = $scriptRoot + "\configuration.xml"
    }
    
    $products = Get-ODTProductToAdd -TargetFilePath $TargetFilePath 
    $addNode = Get-ODTAdd -TargetFilePath $TargetFilePath 

    $sourcePath = $addNode.SourcePath
    $version = $addNode.Version
    $edition = $addNode.OfficeClientEdition

    foreach ($product in $products)
    {
        if ($product) {
          $languages = getProductLanguages -Product $product 
          $existingLangs = checkForLanguagesInSourceFiles -Languages $languages -SourcePath $sourcePath -Version $version -Edition $edition
          if ($product.ProductId) {
              Set-ODTProductToAdd -TargetFilePath $TargetFilePath -ProductId $product.ProductId -LanguageIds $existingLangs | Out-Null
          }
        }
    }

    $localPath = "$env:TEMP\setup.exe"

    Copy-Item -Path $officeCtrPath -Destination $localPath -Force

    $cmdLine = $localPath
    $cmdArgs = "/configure " + '"' + $TargetFilePath + '"'

    Write-Host "Installing Office Click-To-Run..."
	#write log
    $lineNum = Get-CurrentLineNumber    
    $filName = Get-CurrentFileName 
    WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Installing Office Click-To-Run..."
	
    if ($WaitForInstallToFinish) {
        StartProcess -execFilePath $cmdLine -execParams $cmdArgs -WaitForExit $false

        Start-Sleep -Seconds 5

        Wait-ForOfficeCTRInstall -OfficeVersion $OfficeVersion
    }else {
        StartProcess -execFilePath $cmdLine -execParams $cmdArgs -WaitForExit $true
    }
  
    if(($PinToStartMenu) -or ($PinToTaskbar)){
        Write-Host ""

        $ClickToRun = Get-OfficeVersion
        if($ClickToRun.GetType().Name -eq "Object[]"){
            $C2RVersion = $ClickToRun[0]
        } else {
            $C2RVersion = $ClickToRun
        }
            
        $ClickToRun = $true
        $InstallPath = $C2RVersion.InstallPath
        $officeVersionInt = $C2RVersion.Version.Split('.')[0]
      
        if($PinToStartMenu){
            if($PinToStartMenu -eq 'AllOfficeApps'){
                $OfficeAppPinnedStatus = GetOfficeAppVerbStatus
            } else {
                if($PinToStartMenu -ne "None"){
                    $OfficeAppPinnedStatus = GetOfficeAppVerbStatus -OfficeApps $PinToStartMenu
                }
            }
            
            if($OfficeAppPinnedStatus -ne $NULL){
                foreach($app in $OfficeAppPinnedStatus){
                    if($app.PinToStartMenuAvailable -eq $true){
                        Set-OfficePinnedApplication -Action PinToStartMenu -OfficeApps $app.Name -ClickToRun $ClickToRun -InstallPath $InstallPath -OfficeVersion $officeVersionInt
                    }   
                }
            }   
            
            $allPinnedApps = GetOfficeAppVerbStatus

            if($PinToStartMenu -ne 'AllOfficeApps'){
                foreach($app in $allPinnedApps){
                    if($PinToStartMenu -notcontains $app.Name){
                        if($app.PinToStartMenuAvailable -eq $false){
                            Set-OfficePinnedApplication -Action UnpinFromStartMenu -OfficeApps $app.Name -ClickToRun $ClickToRun -InstallPath $InstallPath -OfficeVersion $officeVersionInt
                        }  
                    }
                } 
            }       
        } else {
            if([Environment]::OSVersion.Version.Major -ge 10){
                $OfficeAppPinnedStatus = GetOfficeAppVerbStatus | ? {$_.PinToStartMenuAvailable -eq $true}
                foreach($app in $PinnedStartMenuApps){
                    if($OfficeAppPinnedStatus.Name -contains $app.Name){
                        Set-OfficePinnedApplication -Action PinToStartMenu -OfficeApps $app.Name -ClickToRun $ClickToRun -InstallPath $InstallPath -OfficeVersion $officeVersionInt
                    }
                }   
            }
        }

        if(($PinToTaskbar) -and ([Environment]::OSVersion.Version.Major -lt 10)){
            if($PinToTaskbar -eq 'AllOfficeApps'){
                $OfficeAppPinnedStatus = GetOfficeAppVerbStatus
            } else {
                $OfficeAppPinnedStatus = GetOfficeAppVerbStatus -OfficeApps $PinToTaskbar
            }    
            
            foreach($app in $OfficeAppPinnedStatus){
                if($app.PinToTaskbarAvailable -eq $true){
                    Set-OfficePinnedApplication -Action PinToTaskbar -OfficeApps $app.Name -ClickToRun $ClickToRun -InstallPath $InstallPath -OfficeVersion $officeVersionInt
                    $pinnedApp += $app.Name
                }     
            }
            
            $allPinnedApps = GetOfficeAppVerbStatus

            if($PinToTaskbar -ne 'AllOfficeApps'){
                foreach($app in $allPinnedApps){
                    if($PinToTaskbar -notcontains $app.Name){
                        if($app.PinToTaskbarAvailable -eq $false){
                            Set-OfficePinnedApplication -Action UnpinFromTaskbar -OfficeApps $app.Name -ClickToRun $ClickToRun -InstallPath $InstallPath -OfficeVersion $officeVersionInt
                        }  
                    }
                } 
            }             
        } else {
            if([Environment]::OSVersion.Version.Major -ge 10){
                $OfficeAppPinnedStatus = GetOfficeAppVerbStatus | ? {$_.PinToTaskbarAvailable -eq $true}
                foreach($app in $PinnedTaskbarApps){
                    if($OfficeAppPinnedStatus.Name -contains $app.Name){
                        Set-OfficePinnedApplication -Action PinToTaskbar -OfficeApps $app.Name -ClickToRun $ClickToRun -InstallPath $InstallPath -OfficeVersion $officeVersionInt
                    }
                }
            }
        }
    }
    
    if($InstallProofingTools -eq $true){
        Write-Host ""
        Write-Host "Installing Proofing Tools..."

        if((Get-OfficeVersion).Bitness -eq "32-bit"){
            $proofingToolFileName = "proofingtools2016_en-us-x86.exe"
        } else {
            $proofingToolFileName = "proofingtools2016_en-us-x64.exe"
        }

        $clientCulture = (Get-OfficeVersion).ClientCulture
        $proofingLangLCID = ([globalization.cultureinfo]::GetCultures("allCultures") | where {$_.Name.ToLower() -match $clientCulture}).LCID

        $commandArgs = "/lang:$proofingLangLCID /quiet /passive /norestart"

        Start-Process -FilePath .\$proofingToolFileName -ArgumentList $commandArgs
          
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
 $MSexceptionList = "mui","visio","project","proofing","visual"

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
                try{
                $installReg = "^" + $installPath.Replace('\', '\\')
                $installReg = $installReg.Replace('(', '\(')
                $installReg = $installReg.Replace(')', '\)')
                if ($officeInstallPath -match $installReg) { $officeProduct = $true }
                } catch {}
             }
           }

           if (!$officeProduct) { continue };
           
           $name = $regProv.GetStringValue($HKLM, $path, "DisplayName").sValue          

           $primaryOfficeProduct = $true
           if ($ConfigItemList.Contains($key.ToUpper()) -and $name.ToUpper().Contains("MICROSOFT OFFICE")) {
              foreach($exception in $MSexceptionList){
                 if($name.ToLower() -match $exception.ToLower()){
                    $primaryOfficeProduct = $false
                 }
              }
           } else {
              $primaryOfficeProduct = $false
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

Function checkForLanguagesInSourceFiles() {
    [CmdletBinding()]
    Param(
        [Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true, Position=0)]
        $Languages = $NULL,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string]$SourcePath = $NULL,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string]$Version = $NULL,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string]$Edition = $NULL
    )

    $scriptRoot = GetScriptRoot

    $returnLanguages = @()

    if (!($SourcePath)) {
      $localSource = $scriptRoot + "\Office\Data"
      if (Test-Path -Path $localSource) {
         $SourcePath = $scriptRoot
      }
    }

    if (!($Version)) {
       $localPath = $env:TEMP
       $cabPath = $scriptRoot + "\Office\Data\v$Edition.cab"
       $cabFolderPath = $scriptRoot + "\Office\Data"
       $vdXmlPath = $localPath + "\VersionDescriptor.xml"
       
       if (Test-Path -Path $cabPath) {
          Invoke-Expression -Command "Expand $cabPath -F:VersionDescriptor.xml $localPath" | Out-Null
          $Version = getVersionFromVersionDescriptor -vesionDescriptorPath $vdXmlPath
          Remove-Item -Path $vdXmlPath -Force
       }
    }

    $verionDir = $scriptRoot + "\Office\Data\$Version"
    
    if (Test-Path -Path $verionDir) {
       foreach ($lang in $Languages) {
          $fileName = "stream.x86.$lang.dat"
          if ($Edition -eq "64") {
             $fileName = "stream.x64.$lang.dat"
          }
          
          $langFile = $verionDir + "\" + $fileName 
          
          if (Test-Path -Path $langFile) {
             $returnLanguages += $lang
          }
       }
    }

    return $returnLanguages
}

Function getVersionFromVersionDescriptor() {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [string] $vesionDescriptorPath = $NULL
    )

    [System.XML.XMLDocument]$doc = New-Object System.XML.XMLDocument

    if ($vesionDescriptorPath) {
        $doc.Load($vesionDescriptorPath) | Out-Null
        return $doc.DocumentElement.Available.Build
    }
}

Function getProductLanguages() {
    [CmdletBinding()]
    Param(
        [Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true, Position=0)]
        $Product = $NULL
    )

    $languages = @()

    foreach ($language in $Product.Languages)
    {
      if (!($languages -contains ($language))) {
          $languages += $language
      }
    }

    return $languages
}

Function getUniqueLanguages() {
    [CmdletBinding()]
    Param(
        [Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true, Position=0)]
        $Products = $NULL
    )

    $languages = @()
    foreach ($product in $Products)
    {
       foreach ($language in $product.Languages)
       {
          if (!($languages -contains $language)) {
            $languages += $language
          }
       }
    }

    return $languages
}

Function Get-ODTProductToAdd{
<#
.SYNOPSIS
Gets list of Products and the corresponding language and exlcudeapp values
from the specified configuration file

.PARAMETER All
Switch to return All Products

.PARAMETER ProductId
Id of Product that you want to pull from the configuration file

.PARAMETER TargetFilePath
Required. Full file path for the file.

.Example
Get-ODTProductToAdd -All -TargetFilePath "$env:Public\Documents\config.xml"
Returns all Products and their corresponding Language and Exclude values
if they have them 

.Example
Get-ODTProductToAdd -ProductId "O365ProPlusRetail" -TargetFilePath "$env:Public\Documents\config.xml"
Returns the Product with the O365ProPlusRetail Id and its corresponding
Language and Exclude values

#>
    [CmdletBinding()]
    Param(
        [Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true, Position=0)]
        [string] $ConfigurationXML = $NULL,

        [Microsoft.Office.Products] $ProductId = "Unknown",

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $TargetFilePath,

        [Parameter(ParameterSetName="All")]
        [switch] $All
    )

    Process{
        $TargetFilePath = GetFilePath -TargetFilePath $TargetFilePath

        #Load the file
        [System.XML.XMLDocument]$ConfigFile = New-Object System.XML.XMLDocument

        if ($TargetFilePath) {
           $content = Get-Content $TargetFilePath
           $ConfigFile.LoadXml($content) | Out-Null
        } else {
            if ($ConfigurationXml) 
            {
              $ConfigFile.LoadXml($ConfigurationXml) | Out-Null
              $global:saveLastConfigFile = $NULL
              $global:saveLastFilePath = $NULL
            }
        }

        #Check that the file is properly formatted
        if($ConfigFile.Configuration -eq $null){
            throw $NoConfigurationElement
        }

        if($ConfigFile.Configuration.Add -eq $null){
            throw $NoAddElement
        }

        if($PSCmdlet.ParameterSetName -eq "All"){
            foreach($ProductElement in $ConfigFile.Configuration.Add.Product){
                $Result = New-Object -TypeName PSObject 

                Add-Member -InputObject $Result -MemberType NoteProperty -Name "ProductId" -Value ($ProductElement.GetAttribute("ID"))

                if($ProductElement.Language -ne $null){
                    $ProductLangs = $configfile.Configuration.Add.Product.Language | % {$_.ID}
                    Add-Member -InputObject $Result -MemberType NoteProperty -Name "Languages" -Value $ProductLangs
                    #Add-Member -InputObject $Result -MemberType NoteProperty -Name "Languages" -Value ($ProductElement.Language.GetAttribute("ID"))
                }

                if($ProductElement.ExcludeApp -ne $null){
                    $ProductExlApps = $configfile.Configuration.Add.Product.ExcludeApp | % {$_.ID}
                    Add-Member -InputObject $Result -MemberType NoteProperty -Name "ExcludedApps" -Value $ProductExlApps
                    #Add-Member -InputObject $Result -MemberType NoteProperty -Name "ExcludedApps" -Value ($ProductElement.ExcludeApp.GetAttribute("ID"))
                }
                $Result
            }
        }else{
            if ($ProductId) {
            

                [System.XML.XMLElement]$ProductElement = $ConfigFile.Configuration.Add.Product | where { $_.ID -eq $ProductId }
                if ($ProductElement) {
                $tempId = $ProductElement.GetAttribute("ID")
                
                
                $Result = New-Object -TypeName PSObject 
                Add-Member -InputObject $Result -MemberType NoteProperty -Name "ProductId" -Value $tempId 
                if($ProductElement.Language -ne $null){
                    $ProductLangs = $configfile.Configuration.Add.Product.Language | % {$_.ID}
                    Add-Member -InputObject $Result -MemberType NoteProperty -Name "Languages" -Value $ProductLangs
                    #Add-Member -InputObject $Result -MemberType NoteProperty -Name "Languages" -Value ($ProductElement.Language.GetAttribute("ID"))
                }

                if($ProductElement.ExcludeApp -ne $null){
                    $ProductExlApps = $configfile.Configuration.Add.Product.ExcludeApp | % {$_.ID}
                    Add-Member -InputObject $Result -MemberType NoteProperty -Name "ExcludedApps" -Value $ProductExlApps
                    #Add-Member -InputObject $Result -MemberType NoteProperty -Name "ExcludedApps" -Value ($ProductElement.ExcludeApp.GetAttribute("ID"))
                }
                $Result
                }
            }
        }

    }

}

Function Get-ODTAdd{
<#
.SYNOPSIS
Gets the value of the Add section in the configuration file

.PARAMETER TargetFilePath
Required. Full file path for the file.

.Example
Get-ODTAdd -TargetFilePath "$env:Public\Documents\config.xml"
Returns the value of the Add section if it exists in the specified
file. 

#>
    Param(

        [Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true, Position=0)]
        [string] $ConfigurationXML = $NULL,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $TargetFilePath

    )

    Process{
        $TargetFilePath = GetFilePath -TargetFilePath $TargetFilePath

        #Load the file
        [System.XML.XMLDocument]$ConfigFile = New-Object System.XML.XMLDocument

        if ($TargetFilePath) {
           $content = Get-Content $TargetFilePath
           $ConfigFile.LoadXml($content) | Out-Null
        } else {
            if ($ConfigurationXml) 
            {
              $ConfigFile.LoadXml($ConfigurationXml) | Out-Null
              $global:saveLastConfigFile = $NULL
              $global:saveLastFilePath = $NULL
            }
        }

        #Check that the file is properly formatted
        if($ConfigFile.Configuration -eq $null){
            throw $NoConfigurationElement
        }
        
        $ConfigFile.Configuration.GetElementsByTagName("Add") | Select OfficeClientEdition, SourcePath, Version, Channel, Branch
    }

}

Function Set-ODTDisplay{
<#
.SYNOPSIS
Modifies an existing configuration xml file to set display level and acceptance of the EULA

.PARAMETER Level
Optional. Determines the user interface that the user sees when the 
operation is performed. If Level is set to None, the user sees no UI. 
No progress UI, completion screen, error dialog boxes, or first run 
automatic start UI are displayed. If Level is set to Full, the user 
sees the normal Click-to-Run user interface: Automatic start, 
application splash screen, and error dialog boxes.

.PARAMETER AcceptEULA
If this attribute is set to TRUE, the user does not see a Microsoft 
Software License Terms dialog box. If this attribute is set to FALSE 
or is not set, the user may see a Microsoft Software License Terms dialog box.

.PARAMETER TargetFilePath
Full file path for the file to be modified and be output to.

.Example
Set-ODTLogging -Level "Full" -TargetFilePath "$env:Public/Documents/config.xml"
Sets config show the UI during install

.Example
Set-ODTDisplay -Level "none" -AcceptEULA "True" -TargetFilePath "$env:Public/Documents/config.xml"
Sets config to hide UI and automatically accept EULA during install

.Notes
Here is what the portion of configuration file looks like when modified by this function:

<Configuration>
  ...
  <Display Level="None" AcceptEULA="TRUE" />
  ...
</Configuration>

#>
    Param(

        [Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true, Position=0)]
        [string] $ConfigurationXML = $NULL,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [LogLevel] $Level,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [bool] $AcceptEULA = $true,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $TargetFilePath

    )

    Process{
        $TargetFilePath = GetFilePath -TargetFilePath $TargetFilePath

        #Load file
        [System.XML.XMLDocument]$ConfigFile = New-Object System.XML.XMLDocument

        if ($TargetFilePath) {
           $content = Get-Content $TargetFilePath
           $ConfigFile.LoadXml($content) | Out-Null
        } else {
            if ($ConfigurationXml) 
            {
              $ConfigFile.LoadXml($ConfigurationXml) | Out-Null
              $global:saveLastConfigFile = $NULL
              $global:saveLastFilePath = $NULL
            }
        }

        $global:saveLastConfigFile = $ConfigFile.OuterXml

        #Check for proper root element
        if($ConfigFile.Configuration -eq $null){
            throw $NoConfigurationElement
        }

        #Get display element if it exists
        [System.XML.XMLElement]$DisplayElement = $ConfigFile.Configuration.GetElementsByTagName("Display").Item(0)
        if($ConfigFile.Configuration.Display -eq $null){
            [System.XML.XMLElement]$DisplayElement=$ConfigFile.CreateElement("Display")
            $ConfigFile.Configuration.appendChild($DisplayElement) | Out-Null
        }

        #Set values
        if($Level -ne $null){
            $DisplayElement.SetAttribute("Level", $Level) | Out-Null
        } else {
            if ($PSBoundParameters.ContainsKey('Level')) {
                $ConfigFile.Configuration.Add.RemoveAttribute("Level")
            }
        }

        if($AcceptEULA -ne $null){
            $DisplayElement.SetAttribute("AcceptEULA", $AcceptEULA.ToString().ToUpper()) | Out-Null
        } else {
            if ($PSBoundParameters.ContainsKey('AcceptEULA')) {
                $ConfigFile.Configuration.Add.RemoveAttribute("AcceptEULA")
            }
        }

        $ConfigFile.Save($TargetFilePath) | Out-Null
        $global:saveLastFilePath = $TargetFilePath

        if (($PSCmdlet.MyInvocation.PipelineLength -eq 1) -or `
            ($PSCmdlet.MyInvocation.PipelineLength -eq $PSCmdlet.MyInvocation.PipelinePosition)) {
            Write-Host

            Format-XML ([xml](cat $TargetFilePath)) -indent 4

            Write-Host
            Write-Host "The Office XML Configuration file has been saved to: $TargetFilePath"
        } else {
            $results = new-object PSObject[] 0;
            $Result = New-Object -TypeName PSObject 
            Add-Member -InputObject $Result -MemberType NoteProperty -Name "TargetFilePath" -Value $TargetFilePath
            Add-Member -InputObject $Result -MemberType NoteProperty -Name "Level" -Value $Level
            Add-Member -InputObject $Result -MemberType NoteProperty -Name "AcceptEULA" -Value $AcceptEULA
            $Result
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
    
   $locationPath = (Get-Location).Path
    
    if (!($TargetFilePath.IndexOf('\') -gt -1)) {
      $TargetFilePath = $locationPath + "\" + $TargetFilePath
    }

    return $TargetFilePath
}

Function Get-OfficeCTRRegPath() {
    $path15 = 'SOFTWARE\Microsoft\Office\15.0\ClickToRun'
    $path16 = 'SOFTWARE\Microsoft\Office\ClickToRun'

    if (Test-Path "HKLM:\$path15") {
      return $path15
    } else {
      if (Test-Path "HKLM:\$path16") {
         return $path16
      }
    }
}

Function Set-ODTProductToAdd{
<#
.SYNOPSIS
Modifies an existing configuration xml file to modify a existing product item.

.PARAMETER ExcludeApps
Array of IDs of Apps to exclude from install

.PARAMETER ProductId
Required. ID must be set to a valid ProductRelease ID.
See https://support.microsoft.com/en-us/kb/2842297 for valid ids.

.PARAMETER LanguageIds
Possible values match 'll-cc' pattern (Microsoft Language ids)
The ID value can be set to a valid Office culture language (such as en-us 
for English US or ja-jp for Japanese). The ll-cc value is the language 
identifier.

.PARAMETER TargetFilePath
Full file path for the file to be modified and be output to.

.Example
Add-ODTProductToAdd -ProductId "O365ProPlusRetail" -LanguageId ("en-US", "es-es") -TargetFilePath "$env:Public/Documents/config.xml" -ExcludeApps ("Access", "InfoPath")
Sets config to add the English and Spanish version of office 365 ProPlus
excluding Access and InfoPath

.Example
Add-ODTProductToAdd -ProductId "O365ProPlusRetail" -LanguageId ("en-US", "es-es) -TargetFilePath "$env:Public/Documents/config.xml"
Sets config to add the English and Spanish version of office 365 ProPlus

.Notes
Here is what the portion of configuration file looks like when modified by this function:

<Configuration>
  <Add OfficeClientEdition="64" >
    <Product ID="O365ProPlusRetail">
      <Language ID="en-US" />
      <Language ID="es-es" />
      <ExcludeApp ID="Access">
      <ExcludeApp ID="InfoPath">
    </Product>
  </Add>
  ...
</Configuration>

#>
    [CmdletBinding()]
    Param(
        [Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true, Position=0)]
        [string] $ConfigurationXML = $NULL,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $TargetFilePath = $NULL,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [Microsoft.Office.Products] $ProductId = "Unknown",

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [Alias("LanguageId")]
        [string[]] $LanguageIds = $NULL,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string[]] $ExcludeApps = $NULL

    )

    Process{
        $TargetFilePath = GetFilePath -TargetFilePath $TargetFilePath

        if ($ProductId -eq "Unknown") {
           $ProductId = SelectProductId
        }

        $ProductId = IsValidProductId -ProductId $ProductId
        
        $langCount = $LanguageIds.Count

        if ($langCount -gt 0) {
           foreach ($language in $LanguageIds) {
              $language = IsSupportedLanguage -Language $language
           }
        }

        #Load the file
        [System.XML.XMLDocument]$ConfigFile = New-Object System.XML.XMLDocument
        
        if ($TargetFilePath) {
           $content = Get-Content $TargetFilePath
           $ConfigFile.LoadXml($content) | Out-Null
        } else {
            if ($ConfigurationXml) 
            {
              $ConfigFile.LoadXml($ConfigurationXml) | Out-Null
              $global:saveLastConfigFile = $NULL
              $global:saveLastFilePath = $NULL
              $TargetFilePath = $NULL
            }
        }

        $global:saveLastConfigFile = $ConfigFile.OuterXml

        #Check that the file is properly formatted
        if($ConfigFile.Configuration -eq $null){
            throw $NoConfigurationElement
        }

        [System.XML.XMLElement]$AddElement=$NULL
        if($ConfigFile.Configuration.Add -eq $null){
           throw "Cannot find 'Add' element"
        }

        $AddElement = $ConfigFile.Configuration.Add 

        #Set the desired values
        [System.XML.XMLElement]$ProductElement = $ConfigFile.Configuration.Add.Product | Where { $_.ID -eq $ProductId }
        if($ProductElement -eq $null){
           throw "Cannot find Product with Id '$ProductId'"
        }

        if ($LanguageIds) {
            $existingLangs = $ProductElement.selectnodes("./Language")
            if ($existingLangs.count -gt 0) {
                foreach ($lang in $existingLangs) {
                  $ProductElement.removeChild($lang) | Out-Null
                }

                foreach($LanguageId in $LanguageIds){
                    [System.XML.XMLElement]$LanguageElement = $ProductElement.Language | Where { $_.ID -eq $LanguageId }
                    if($LanguageElement -eq $null){
                        [System.XML.XMLElement]$LanguageElement=$ConfigFile.CreateElement("Language")
                        $ProductElement.appendChild($LanguageElement) | Out-Null
                        $LanguageElement.SetAttribute("ID", $LanguageId) | Out-Null
                    }
                }
            }
        }

        if ($ExcludeApps) {
            $existingExcludes = $ProductElement.selectnodes("./ExcludeApp")
            if ($existingExcludes.count -gt 0) {
                foreach ($exclude in $existingLangs) {
                  $ProductElement.removeChild($exclude) | Out-Null
                }
            }

            foreach($ExcludeApp in $ExcludeApps){
                [System.XML.XMLElement]$ExcludeAppElement = $ProductElement.ExcludeApp | Where { $_.ID -eq $ExcludeApp }
                if($ExcludeAppElement -eq $null){
                    [System.XML.XMLElement]$ExcludeAppElement=$ConfigFile.CreateElement("ExcludeApp")
                    $ProductElement.appendChild($ExcludeAppElement) | Out-Null
                    $ExcludeAppElement.SetAttribute("ID", $ExcludeApp) | Out-Null
                }
            }
        }

        $ConfigFile.Save($TargetFilePath) | Out-Null
        $global:saveLastFilePath = $TargetFilePath

        if (($PSCmdlet.MyInvocation.PipelineLength -eq 1) -or `
            ($PSCmdlet.MyInvocation.PipelineLength -eq $PSCmdlet.MyInvocation.PipelinePosition)) {
            Write-Host

            Format-XML ([xml](cat $TargetFilePath)) -indent 4

            Write-Host
            Write-Host "The Office XML Configuration file has been saved to: $TargetFilePath"
        } else {
            $results = new-object PSObject[] 0;
            $Result = New-Object -TypeName PSObject 
            Add-Member -InputObject $Result -MemberType NoteProperty -Name "TargetFilePath" -Value $TargetFilePath
            Add-Member -InputObject $Result -MemberType NoteProperty -Name "ProductId" -Value $ProductId
            Add-Member -InputObject $Result -MemberType NoteProperty -Name "LanguageIds" -Value $LanguageIds
            $Result
        }


    }

}

Function Wait-ForOfficeCTRInstall() {
    [CmdletBinding()]
    Param(
        [Parameter()]
        [int] $TimeOutInMinutes = 120,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [OfficeCTRVersion] $OfficeVersion = "Office2016"
    )

    begin {
        $HKLM = [UInt32] "0x80000002"
        $HKCR = [UInt32] "0x80000000"
    }

    process {
        Write-Host "Waiting for Install to Begin..."
 
        #Start-Sleep -Seconds 25

        if($OfficeVersion -eq 'Office2016'){
            $mainRegPath = 'SOFTWARE\Microsoft\Office\ClickToRun'
        } else {
            $mainRegPath = Get-OfficeCTRRegPath
        } 

        $scenarioPath = $mainRegPath + "\scenario"

        $regProv = Get-Wmiobject -list "StdRegProv" -namespace root\default -ErrorAction Stop

        [DateTime]$startTime = Get-Date

        [string]$executingScenario = ""
        $failure = $false
        $updateRunning=$false
        [string[]]$trackProgress = @()
        [string[]]$trackComplete = @()
        
        $timeout = New-TimeSpan -Minutes 2
        $sw = [diagnostics.stopwatch]::StartNew()
        while ($sw.elapsed -lt $timeout){
            try {
                $exScenario = $regProv.GetStringValue($HKLM, $mainRegPath, "ExecutingScenario")
                if($exScenario.sValue){ break; }
            } catch {}

            Start-Sleep -Seconds 5
        }
       
        if ($exScenario) {
            $executingScenario = $exScenario.sValue
        }
         
        do {
            $allComplete = $true
            $scenarioKeys = $regProv.EnumKey($HKLM, $scenarioPath)
            foreach ($scenarioKey in $scenarioKeys.sNames) {
                if (!($executingScenario)) { continue }
                if ($scenarioKey.ToLower() -eq $executingScenario.ToLower()) {
                    $taskKeyPath = $scenarioPath + "\$scenarioKey\TasksState"
                    $taskValues = $regProv.EnumValues($HKLM, $taskKeyPath).sNames

                    foreach ($taskValue in $taskValues) {
                        [string]$status = $regProv.GetStringValue($HKLM, $taskKeyPath, $taskValue).sValue
                        $operation = $taskValue.Split(':')[0]
                        $keyValue = $taskValue

                        if ($status.ToUpper() -eq "TASKSTATE_FAILED") {
                            $failure = $true
                        }

                        $displayValue = showTaskStatus -Operation $operation -Status $status -DateTime (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')

                        if (($status.ToUpper() -eq "TASKSTATE_COMPLETED") -or`
                            ($status.ToUpper() -eq "TASKSTATE_CANCELLED") -or`
                            ($status.ToUpper() -eq "TASKSTATE_FAILED")) {
                                if (($trackProgress -contains $keyValue) -and !($trackComplete -contains $keyValue)) {
                                    $displayValue
                                    $trackComplete += $keyValue
                                    Start-Sleep -Seconds 1
                                }
                        } else {
                            $allComplete = $false
                            $updateRunning = $true

                            if ($trackProgress -notcontains $keyValue) {
                                $displayValue
                                $trackProgress += $keyValue                                
                                Start-Sleep -Seconds 1
                            }
                        }
                    }
                }
            }

            if ($startTime -lt (Get-Date).AddHours(-$TimeOutInMinutes)) {
                throw "Waiting for Update Timed-Out"
                break;
            }

            if($allComplete){
                $updateRunning = $false
            }

            Start-Sleep -Seconds 5

        } while($updateRunning -eq $true)
    
        if($failure){
            Write-Host ""
            Write-Host 'Update failed'
        } else {
            if($trackProgress.Count -gt 0){
                Write-Host ""
                Write-Host 'Update complete'
            } else {
                Write-Host ""
                Write-Host 'Update not running'
            }
        } 
    }
}

function showTaskStatus() {
    [CmdletBinding()]
    Param(
        [Parameter()]
        [string] $Operation = "",

        [Parameter()]
        [string] $Status = "",

        [Parameter()]
        [string] $DateTime = ""
    )

    $Result = New-Object -TypeName PSObject 
    Add-Member -InputObject $Result -MemberType NoteProperty -Name "Operation" -Value $Operation
    Add-Member -InputObject $Result -MemberType NoteProperty -Name "Status" -Value $Status
    Add-Member -InputObject $Result -MemberType NoteProperty -Name "DateTime" -Value $DateTime
    return $Result
}

Function StartProcess {
	Param
	(
        [Parameter()]
		[String]$execFilePath,

        [Parameter()]
        [String]$execParams,

        [Parameter()]
        [bool]$WaitForExit = $false


	)

    Try
    {
        $startExe = new-object System.Diagnostics.ProcessStartInfo
        $startExe.FileName = $execFilePath
        $startExe.Arguments = $execParams
        $startExe.CreateNoWindow = $false
        $startExe.UseShellExecute = $false

        $execStatement = [System.Diagnostics.Process]::Start($startExe) 
        if ($WaitForExit) {
           $execStatement.WaitForExit()
        }
    }
    Catch
    {
        Write-Log -Message $_.Exception.Message -severity 1 -component "Office 365 Update Anywhere"
    }
}

Function GetScriptRoot() {
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

Function Format-XML ([xml]$xml, $indent=2) { 
    $StringWriter = New-Object System.IO.StringWriter 
    $XmlWriter = New-Object System.XMl.XmlTextWriter $StringWriter 
    $xmlWriter.Formatting = "indented" 
    $xmlWriter.Indentation = $Indent 
    $xml.WriteContentTo($XmlWriter) 
    $XmlWriter.Flush() 
    $StringWriter.Flush() 
    Write-Output $StringWriter.ToString() 
}

function Set-OfficePinnedApplication { 
<#  
.SYNOPSIS  
Automate the process for pinning or unpinning Office apps

.DESCRIPTION  
Pin or unpin Office apps from the Start Menu or Taskbarb setting the action

.EXAMPLE 
Set-PinnedApplication -Action PinToTaskbar

.EXAMPLE 
Set-PinnedApplication -Action UnPinFromTaskbar 

.EXAMPLE 
Set-PinnedApplication -Action PinToStartMenu

.EXAMPLE 
Set-PinnedApplication -Action UnPinFromStartMenu 

#>  
    [CmdletBinding()] 
    param( 
        [Parameter(Mandatory=$true)]
        [PinAction]$Action,

        [Parameter()]
        [ValidateSet("Word","Excel","PowerPoint","OneNote","Access","Publisher","Outlook","Skype for Business",
                     "OneDrive for Business","Project","Visio")]
        [string[]]$OfficeApps = $null,

        [Parameter()]
        [string]$ClickToRun,

        [Parameter()]
        [string]$InstallPath,

        [Parameter()]
        [string]$OfficeVersion
    )

    if(!$ClickToRun){
        $ctr = Get-OfficeVersion
        if($ctr.GetType().Name -eq "Object[]"){
            $ClickToRun = $ctr[0].ClickToRun
        } else {
            $ClickToRun = (Get-OfficeVersion).ClickToRun
        }
    }
    
    if(!$InstallPath){
        $ctr = Get-OfficeVersion
        if($ctr.GetType().Name -eq "Object[]"){
            $InstallPath = $ctr[0].InstallPath
        } else {
            $InstallPath = (Get-OfficeVersion).InstallPath
        }   
    }
    
    if(!$officeVersion){
        $ctr = Get-OfficeVersion
        if($ctr.GetType().Name -eq "Object[]"){
            $officeVersion = $ctr[0].Version.Split('.')[0]
        } else {
            $officeVersion = (Get-OfficeVersion).Version.Split('.')[0]
        }             
    }

    if($InstallPath.GetType().Name -eq "Object[]"){
        $InstallPath = $InstallPath[0]
    }

    if($ClickToRun -eq $true) {
        $officeAppPath = $InstallPath + "\root\Office" + $officeVersion
    } else {
        $officeAppPath = $InstallPath + "Office" + $officeVersion
    }

    $officeAppList = @()

    if(!$OfficeApps){
        $officeAppList = @("WINWORD.EXE", "EXCEL.EXE", "POWERPNT.EXE", "ONENOTE.EXE", "MSACCESS.EXE", "MSPUB.EXE", "OUTLOOK.EXE",
                           "lync.exe", "GROOVE.EXE", "WINPROJ.EXE", "VISIO.EXE")
    } else {
        foreach($app in $OfficeApps){
            switch($app){
                "Word" {
                    $officeAppList += "WINWORD.EXE"
                }
                "Excel" {
                    $officeAppList += "EXCEL.EXE"
                }
                "PowerPoint" {
                    $officeAppList += "POWERPNT.EXE"
                }
                "OneNote" {
                    $officeAppList += "ONENOTE.EXE"
                }
                "Access" {
                    $officeAppList += "MSACCESS.EXE"
                }
                "Publisher" {
                    $officeAppList += "MSPUB.EXE"
                }
                "Outlook" {
                    $officeAppList += "OUTLOOK.EXE"
                }
                "Skype for Business" {
                    $officeAppList += "lync.exe"
                }
                "OneDrive for Business" {
                    $officeAppList += "GROOVE.EXE"
                }
                "Project" {
                    $officeAppList += "WINPROJ.EXE"
                }
                "Visio" {
                    $officeAppList += "VISIO.EXE"
                }
            }
        }
    }

    foreach($app in $officeAppList){
        if(Test-Path ($officeAppPath + "\$app")){
            switch($Action) {
                "PinToStartMenu" {
                    Write-Host "Pinning $app to the Start Menu..."
                    if([Environment]::OSVersion.Version.Major -ge 10){
                        $actionId = '51201'
                    } else { 
                        $actionId = '5381'
                    }
                }
                "UnpinFromStartMenu" {
                    Write-Host "Removing $app from the Start Menu..."
                    if([Environment]::OSVersion.Version.Major -ge 10){
                        $actionId = '51394'
                    } else { 
                        $actionId = '5382'
                    } 
                }
                "PinToTaskbar" {
                    Write-Host "Pinning $app to the TaskBar..."
                    if([Environment]::OSVersion.Version.Major -ge 10){
                        throw "Unable to pin items to the taskbar in Windows 10"
                    }
            
                    $actionId = '5386'
                }
                "UnpinFromTaskbar" {
                    Write-Host "Removing $app from the TaskBar..."
                    $actionId = '5387'
                }
            }

            InvokeVerb -FilePath ($officeAppPath + "\$app") -Verb $(GetVerb -VerbId $actionId) -officeVersion $officeVersion
        }
    } 
} 

function GetVerb { 
    param(
        [int]$verbId
    ) 

    try { 
        $t = [type]"CosmosKey.Util.MuiHelper" 
    } catch { 
        $def = [Text.StringBuilder]"" 
        [void]$def.AppendLine('[DllImport("user32.dll")]') 
        [void]$def.AppendLine('public static extern int LoadString(IntPtr h,uint id, System.Text.StringBuilder sb,int maxBuffer);') 
        [void]$def.AppendLine('[DllImport("kernel32.dll")]') 
        [void]$def.AppendLine('public static extern IntPtr LoadLibrary(string s);') 
        Add-Type -MemberDefinition $def.ToString() -name MuiHelper -namespace CosmosKey.Util             
    } 
    if($global:CosmosKey_Utils_MuiHelper_Shell32 -eq $null){         
        $global:CosmosKey_Utils_MuiHelper_Shell32 = [CosmosKey.Util.MuiHelper]::LoadLibrary("shell32.dll") 
    } 
    $maxVerbLength=255 
    $verbBuilder = new-object Text.StringBuilder "",$maxVerbLength 
    [void][CosmosKey.Util.MuiHelper]::LoadString($CosmosKey_Utils_MuiHelper_Shell32,$verbId,$verbBuilder,$maxVerbLength) 
    
    return $verbBuilder.ToString() 
} 

function InvokeVerb { 
    param([string]$FilePath,$verb,$officeVersion) 

    $verb = $verb.Replace("&","") 
    $path= split-path $FilePath 
    $shell=new-object -com "Shell.Application"  
    $folder=$shell.Namespace($path)    
    $item = $folder.Parsename((split-path $FilePath -leaf)) 
    $itemVerb = $item.Verbs() | ? {$_.Name.Replace("&","") -eq $verb} 
   
    try{
        if(([Environment]::OSVersion.Version.Major -ge 10) -and ($verb -eq 'Unpin from taskbar')){
            Remove-PinnedOfficeAppsForWindows10 -OfficeApp $item.Name -officeVersion $officeVersion
        }else { 
            $itemVerb.DoIt() 
        }
    }catch{}
     
}

function Remove-PinnedOfficeAppsForWindows10() {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$OfficeApp,

        [Parameter()]
        [string]$officeVersion
    )

    $Action = 'Unpin from taskbar'

    switch($OfficeApp){
        "WINWORD" {
            $officeAppName = "Word"  
        }
        "EXCEL" {
            $officeAppName = "Excel"
        }
        "POWERPNT" {
            $officeAppName = "PowerPoint"
        }
        "ONENOTE" {
            $officeAppName = "OneNote"
        }
        "MSACCESS" {
            $officeAppName = "Access"
        }
        "MSPUB" { 
            $officeAppName = "Publisher"
        }
        "OUTLOOK" {
            $officeAppName = "Outlook"
        }
        "lync" {
            $officeAppName = "Skype for Business"
        }
        "GROOVE" {
            $officeAppName = "OneDrive for Business"
        }
        "WINPROJ" {
            $officeAppName = "Project"
        }
        "VISIO" {
            $officeAppName = "Visio"
        }
    }

        switch($officeVersion){
        "11" {
            $officeAppVersion = "Microsoft Office " + $officeAppName + " 2003"
        }
        "12" {
            $officeAppVersion = "Microsoft Office " + $officeAppName + " 2007"
        }
        "14" {
            $officeAppVersion = "Microsoft " + $officeAppName + " 2010"
        }
        "15" {
            $officeAppVersion = $officeAppName + " 2013"
        }
        "16" {
            $officeAppVersion = $officeAppName + " 2016"
        }
    }

    ((New-Object -Com Shell.Application).NameSpace('shell:::{4234d49b-0245-4df3-b780-3893943456e1}').Items() | ? {$_.Name -like $officeAppVersion}).Verbs() | ? {$_.Name.replace('&','') -match $Action} | % {$_.DoIt()}
       
}

function GetOfficeAppVerbStatus{
    [CmdletBinding()]
    param(
        [Parameter()]
        [ValidateSet("Word","Excel","PowerPoint","OneNote","Access","Publisher","Outlook","Skype for Business",
                     "OneDrive for Business","Project","Visio")]
        [string[]]$OfficeApps
    )

    Begin{
        $defaultDisplaySet = 'Name','PinToStartMenuAvailable', 'PinToTaskbarAvailable'
        $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet',[string[]]$defaultDisplaySet)
        $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
    }

    Process{
        $results = new-object PSObject[] 0;

        $ctr = Get-OfficeVersion 
                     
        if($ctr -ne $null){
            if($ctr.GetType().Name -eq "Object[]"){
                $ctr = $ctr[0]
                $officeversion = $ctr.Version.Split('.')[0]                      
            } else {
                $officeVersion = (Get-OfficeVersion).Version.Split('.')[0]
            }
        }
        
        $InstallPath = $ctr.InstallPath 
        $ctr = $ctr.ClickToRun

        if($ctr -eq $true) {
            $officeAppPath = $InstallPath + "\root\Office" + $officeVersion
        } else {
            $officeAppPath = $InstallPath + "Office" + $officeVersion
        }

        $availableApps = @()
        $officeAppList = @()

        if(!$OfficeApps){
            $officeAppList += @("WINWORD.EXE", "EXCEL.EXE", "POWERPNT.EXE", "ONENOTE.EXE", "MSACCESS.EXE", "MSPUB.EXE", "OUTLOOK.EXE",
                            "lync.exe", "GROOVE.EXE", "WINPROJ.EXE", "VISIO.EXE")
        } else {
            foreach($app in $OfficeApps){
                switch($app){
                    "Word" {
                        $officeAppList += "WINWORD.EXE"
                    }
                    "Excel" {
                        $officeAppList += "EXCEL.EXE"
                    }
                    "PowerPoint" {
                        $officeAppList += "POWERPNT.EXE"
                    }
                    "OneNote" {
                        $officeAppList += "ONENOTE.EXE"
                    }
                    "Access" {
                        $officeAppList += "MSACCESS.EXE"
                    }
                    "Publisher" {
                        $officeAppList += "MSPUB.EXE"
                    }
                    "Outlook" {
                        $officeAppList += "OUTLOOK.EXE"
                    }
                    "Lync" {
                        $officeAppList += "lync.exe"
                    }
                    "OneDriveForBusiness" {
                        $officeAppList += "GROOVE.EXE"
                    }
                    "Project" {
                        $officeAppList += "WINPROJ.EXE"
                    }
                    "Visio" {
                        $officeAppList += "VISIO.EXE"
                    }
                }
            }
        }

        foreach($app in $OfficeAppList){
            if(Test-Path ($officeAppPath + "\$app")){
                $availableApps += $app
            }
        }

        foreach($app in $availableApps){
            switch($app){
                "WINWORD.EXE" {
                    $officeAppName = "Word"  
                }
                "EXCEL.EXE" {
                    $officeAppName = "Excel"
                }
                "POWERPNT.EXE" {
                    $officeAppName = "PowerPoint"
                }
                "ONENOTE.EXE" {
                    $officeAppName = "OneNote"
                }
                "MSACCESS.EXE" {
                    $officeAppName = "Access"
                }
                "MSPUB.EXE" { 
                    $officeAppName = "Publisher"
                }
                "OUTLOOK.EXE" {
                    $officeAppName = "Outlook"
                }
                "lync.exe" {
                    $officeAppName = "Skype for Business"
                }
                "GROOVE.EXE" {
                    $officeAppName = "OneDrive for Business"
                }
                "WINPROJ.EXE" {
                    $officeAppName = "Project"
                }
                "VISIO.EXE" {
                    $officeAppName = "Visio"
                }
            }

            if(!($officeAppName -eq "OneDrive for Business")){
                switch($officeVersion){
                    "11" {
                        $officeAppVersion = "Microsoft Office " + $officeAppName + " 2003"
                    }
                    "12" {
                        $officeAppVersion = "Microsoft Office " + $officeAppName + " 2007"
                    }
                    "14" {
                        $officeAppVersion = "Microsoft " + $officeAppName + " 2010"
                    }
                    "15" {
                        $officeAppVersion = $officeAppName + " 2013"
                    }
                    "16" {
                        $officeAppVersion = $officeAppName + " 2016"
                    }
                }
            }
            
            [bool]$availablePinToStartMenu = $false
            [bool]$availablePinToTaskbar = $false
            
            if([Environment]::OSVersion.Version.Major -ge 10){
                $verbs = ((New-Object -Com Shell.Application).NameSpace('shell:::{4234d49b-0245-4df3-b780-3893943456e1}').Items() | ? {$_.Name -like $officeAppVersion}).Verbs() | select Name
                
                foreach($verb in $verbs.name){
                    switch($verb.Replace('&','')){
                        "Pin to Start" {
                            $availablePinToStartMenu = $true
                        }
                        "Pin to Taskbar" {
                            $availablePinToTaskbar = $true
                        }
                    }
                }
            } else {
                $pinActions = @("5381","5386")
                
                foreach($action in $pinActions){
                    $verb = GetVerb -verbId $action
                    $verb = $verb.Replace("&","")
                    $FilePath = $officeAppPath + "\" + $app
                    $path = Split-Path $FilePath
                    $shell = New-Object -ComObject "Shell.Application"
                    $folder = $shell.Namespace($path)
                    $item = $folder.Parsename((Split-Path $filepath -Leaf))
                    $itemverb = $item.Verbs() | ? {$_.Name.Replace("&","") -eq $verb}
                    
                    if($itemverb){
                        switch($verb){
                            "Pin to Start Menu" {
                                $availablePinToStartMenu = $true
                            }
                            "Pin to Taskbar" {
                                $availablePinToTaskbar = $true
                            }
                        }
                    }
                }         
            }
            
            $object = New-Object PSObject -Property @{Name = $officeAppName; PinToStartMenuAvailable = $availablePinToStartMenu; PinToTaskbarAvailable = $availablePinToTaskbar;}
                                                      
            $object | Add-Member MemberSet PSStandardMembers $PSStandardMembers
            $results += $object 
        }

        return $results
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
