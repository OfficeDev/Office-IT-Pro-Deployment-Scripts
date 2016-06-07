try {
$enumDef = "
using System;
       [FlagsAttribute]
       public enum OfficeChannel
       {
          FirstReleaseCurrent = 0,
          Current = 1,
          FirstReleaseDeferred = 2,
          Deferred = 3
       }
"
Add-Type -TypeDefinition $enumDef -ErrorAction SilentlyContinue
} catch { }

try {
$enumBitness = "
using System;
       [FlagsAttribute]
       public enum Bitness
       {
          Both = 0,
          v32 = 1,
          v64 = 2
       }
"
Add-Type -TypeDefinition $enumBitness -ErrorAction SilentlyContinue
} catch { }

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
     }
}
"
Add-Type -TypeDefinition $enum -ErrorAction SilentlyContinue
} catch {}

try {
$enum = "
using System;
 
namespace Microsoft.Office
{
     [FlagsAttribute]
     public enum ProductSelection
     {
         All = 0,
         O365ProPlusRetail = 1,
         O365BusinessRetail = 2,
         VisioProRetail = 4,
         ProjectProRetail = 8,
         SPDRetail = 16,
         VisioProXVolume = 32,
         VisioStdXVolume = 64,
         ProjectProXVolume = 128,
         ProjectStdXVolume = 256,
     }
}
"
Add-Type -TypeDefinition $enum -ErrorAction SilentlyContinue
} catch {}

[System.Collections.ArrayList]$missingFiles = @()

Function Write-Log {
 
    PARAM
	(
         [String]$Message,
         [String]$Path = $Global:UpdateAnywhereLogPath,
         [String]$LogName = $Global:UpdateAnywhereLogFileName,
         [int]$severity,
         [string]$component
	)
 
    try {
        $Path = $Global:UpdateAnywhereLogPath
        $LogName = $Global:UpdateAnywhereLogFileName
        if([String]::IsNullOrWhiteSpace($Path)){
            # Get Windows Folder Path
            $windowsDirectory = [Environment]::GetFolderPath("Windows")

            # Build log folder
            $Path = "$windowsDirectory\CCM\logs"
        }

        if([String]::IsNullOrWhiteSpace($LogName)){
             # Set log file name
            $LogName = "Office365UpdateAnywhere.log"
        }
        # Build log path
        $LogFilePath = Join-Path $Path $LogName

        # Create log file
        If (!($(Test-Path $LogFilePath -PathType Leaf)))
        {
            $null = New-Item -Path $LogFilePath -ItemType File -ErrorAction SilentlyContinue
        }

	    $TimeZoneBias = Get-WmiObject -Query "Select Bias from Win32_TimeZone"
        $Date= Get-Date -Format "HH:mm:ss.fff"
        $Date2= Get-Date -Format "MM-dd-yyyy"
        $type=1
 
        if ($LogFilePath) {
           "<![LOG[$Message]LOG]!><time=$([char]34)$date$($TimeZoneBias.bias)$([char]34) date=$([char]34)$date2$([char]34) component=$([char]34)$component$([char]34) context=$([char]34)$([char]34) type=$([char]34)$severity$([char]34) thread=$([char]34)$([char]34) file=$([char]34)$([char]34)>"| Out-File -FilePath $LogFilePath -Append -NoClobber -Encoding default
        }
    } catch {

    }
}

Function Set-Reg {
	PARAM
	(
        [String]$hive,
        [String]$keyPath,
	    [String]$valueName,
	    [String]$value,
        [String]$Type
    )

    Try
    {
        $null = New-ItemProperty -Path "$($hive):\$($keyPath)" -Name "$($valueName)" -Value "$($value)" -PropertyType $Type -Force -ErrorAction Stop
    }
    Catch
    {
        Write-Log -Message $_.Exception.Message -severity 3 -component $LogFileName
    }
}

Function StartProcess {
	Param
	(
		[String]$execFilePath,
        [String]$execParams
	)

    Try
    {
        $execStatement = [System.Diagnostics.Process]::Start( $execFilePath, $execParams ) 
        $execStatement.WaitForExit()
    }
    Catch
    {
        Write-Log -Message $_.Exception.Message -severity 1 -component "Office 365 Update Anywhere"
    }
}

function Test-ItemPathUNC() {    [CmdletBinding()]	
    Param
	(	    [Parameter(Mandatory=$true)]
	    [String]$Path,	    [Parameter()]
	    [String]$FileName = $null    )    Process {       $pathExists = $false       if ($FileName) {         $filePath = "$Path\$FileName"         $pathExists = [System.IO.File]::Exists($filePath)       } else {         $pathExists = [System.IO.Directory]::Exists($Path)         if (!($pathExists)) {            $pathExists = [System.IO.File]::Exists($Path)         }       }       return $pathExists;    }}

function Copy-ItemUNC() {    [CmdletBinding()]	
    Param
	(	    [Parameter(Mandatory=$true)]
	    [String]$SourcePath,	    [Parameter(Mandatory=$true)]
	    [String]$TargetPath,	    [Parameter(Mandatory=$true)]
	    [String]$FileName    )    Process {       $drvLetter = FindAvailable       $Network = New-Object -ComObject "Wscript.Network"       try {           if (!($drvLetter.EndsWith(":"))) {               $drvLetter += ":"           }           $target = $drvLetter + "\"           $Network.MapNetworkDrive($drvLetter, $TargetPath)                                 #New-PSDrive -Name $drvLetter -PSProvider FileSystem -Root $TargetPath | Out-Null           Copy-Item -Path $SourcePath -Destination $target -Force       } finally {         #Remove-PSDrive $drvLetter         $Network.RemoveNetworkDrive($drvLetter)       }    }}

function FindAvailable() {
   #$drives = Get-PSDrive | select Name
   $drives = Get-WmiObject -Class Win32_LogicalDisk | select DeviceID

   for($n=90;$n -gt 68;$n--) {
      $letter= [char]$n + ":"
      $exists = $drives | where { $_.DeviceID -eq $letter }
      if ($exists) {
        if ($exists.Count -eq 0) {
            return $letter
        }
      } else {
        return $letter
      }
   }
   return $null
}

function Get-XMLLanguages() {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
       	[Parameter(Mandatory=$true)]
	    [String]$Path
    )
    Process {
      [string[]]$languages = @()
      [xml]$configXml = Get-Content $Path
      if ($configXml.Configuration.Add) {
         foreach ($product in $configXml.Configuration.Add.Product) {
             foreach ($language in $product.Language) {
                $languages += $language.ID
             }
         }
      }
      return $languages
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
Version: 1.0.4
DateCreated: 2015-07-01
DateUpdated: 2015-08-28

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

  $results = Get-Unique -InputObject $results 

  return $results;
}

}

Function Get-InstalledLanguages() {
    [CmdletBinding()]
    Param(
    )
    process {
       $returnLangs = @()
       $mainRegPath = Get-OfficeCTRRegPath

       if ($mainRegPath) {
          if (Test-Path -Path "hklm:\$mainRegPath\ProductReleaseIDs") {
               $activeConfig = Get-ItemProperty -Path "hklm:\$mainRegPath\ProductReleaseIDs"
               $activeId = $activeConfig.ActiveConfiguration
               $languages = Get-ChildItem -Path "hklm:\$mainRegPath\ProductReleaseIDs\$activeId\culture"

               foreach ($language in $languages) {
                  $lang = Get-ItemProperty -Path  $language.pspath
                  $keyName = $lang.PSChildName
                  if ($keyName.Contains(".")) {
                      $keyName = $keyName.Split(".")[0]
                  }

                  if ($keyName.ToLower() -ne "x-none") {
                     $culture = New-Object system.globalization.cultureinfo($keyName)
                     $returnLangs += $culture
                  }
               }
          }
       }

       return $returnLangs
    }
}

Function Get-OfficeCDNUrl() {
    $CDNBaseUrl = (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration -Name CDNBaseUrl -ErrorAction SilentlyContinue).CDNBaseUrl
    if (!($CDNBaseUrl)) {
       $CDNBaseUrl = (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration -Name CDNBaseUrl -ErrorAction SilentlyContinue).CDNBaseUrl
    }
    if (!($CDNBaseUrl)) {
        $OfficeRegPath = ""
        $path15 = 'HKLM:\SOFTWARE\Microsoft\Office\15.0\ClickToRun\ProductReleaseIDs\Active\stream'
        $path16 = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\ProductReleaseIDs\Active\stream'
        if (Test-Path -Path $path16) { $OfficeRegPath = $path16 }
        if (Test-Path -Path $path15) { $OfficeRegPath = $path15 }

        if($OfficeRegPath) {
            $items = Get-Item -Path $OfficeRegPath -ErrorAction SilentlyContinue | Select-Object -ExpandProperty property
            if ($items) {
                $properties = $items | ForEach-Object {
                   New-Object psobject -Property @{"property"=$_; "Value" = (Get-ItemProperty -Path . -Name $_).$_}
                }

                $value = $properties | Select Value
                $firstItem = $value[0]
                [string] $cdnPath = $firstItem.Value

                $CDNBaseUrl = Select-String -InputObject $cdnPath -Pattern "http://officecdn.microsoft.com/.*/.{8}-.{4}-.{4}-.{4}-.{12}" -AllMatches | % { $_.Matches } | % { $_.Value }
            }
        }
    }
    return $CDNBaseUrl
}

Function Get-OfficeC2RVersion() {
    $VersionToReport = (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration -Name VersionToReport -ErrorAction SilentlyContinue).VersionToReport
    if (!($VersionToReport)) {
       $VersionToReport = (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration -Name ClientVersionToReport -ErrorAction SilentlyContinue).ClientVersionToReport
    }
    return $VersionToReport
}

Function Get-OfficeCTRRegPath() {
    $path15 = 'SOFTWARE\Microsoft\Office\15.0\ClickToRun'
    $path16 = 'SOFTWARE\Microsoft\Office\ClickToRun'
    if (Test-Path "HKLM:\$path16") {
        return $path16
    }
    else {
        if (Test-Path "HKLM:\$path15") {
            return $path15
        }
    }
}

function Test-URL {
   param( 
      [string]$url = $NULL
   )

   [bool]$validUrl = $false
   try {
     $req = [System.Net.HttpWebRequest]::Create($url);
     $res = $req.GetResponse()

     if($res.StatusCode -eq "OK") {
        $validUrl = $true
     }
     $res.Close(); 
   } catch {
      Write-Host "Invalid UpdateSource. File Not Found: $url" -ForegroundColor Red
      $validUrl = $false
      throw;
   }

   return $validUrl
}

function Change-UpdatePathToChannel {
   [CmdletBinding()]
   param( 
     [Parameter()]
     [string] $UpdatePath,

     [Parameter()]
     [bool] $ValidateUpdateSourceFiles = $true,

     [Parameter()]
     [string] $Channel = $null,

     [Parameter()]
     [string] $Bitness = $null
   )

   $newUpdatePath = $UpdatePath
   $newUpdateLong = $UpdatePath

   if (!($Channel)) {
      $detectedChannel = Detect-Channel 
   }

   if ($detectedChannel) {
       $branchName = $detectedChannel.branch
   } else {
      if ($Channel) {
         $branchName = $Channel
      } else {
         $branchName = "Deferred"
      }
   }

   $branchShortName = "DC"
   if ($branchName.ToLower() -eq "current") {
      $branchShortName = "CC"
   }
   if ($branchName.ToLower() -eq "firstreleasecurrent") {
      $branchShortName = "FRCC"
   }
   if ($branchName.ToLower() -eq "firstreleasedeferred") {
      $branchShortName = "FRDC"
   }
   if ($branchName.ToLower() -eq "deferred") {
      $branchShortName = "DC"
   }

   $channelNames = @("FRCC", "CC", "FRDC", "DC")
   $channelLongNames = @("FirstReleaseCurrent", "Current", "FirstReleaseDeferred", "Deferred", "Business", "FirstReleaseBusiness")

   $madeChange = $false
   foreach ($channelName in $channelNames) {
      if ($UpdatePath.ToUpper().EndsWith("\$channelName")) {
         $newUpdatePath = $newUpdatePath -replace "\\$channelName", "\$branchShortName"
         $newUpdateLong = $newUpdateLong -replace "\\$channelName", "\$branchName"
         $madeChange = $true
      } 
      if ($UpdatePath.ToUpper().Contains("\$channelName\")) {
         $newUpdatePath = $newUpdatePath -replace "\\$channelName\\", "\$branchShortName\"
         $newUpdateLong = $newUpdateLong -replace "\\$channelName\\", "\$branchName\"
         $madeChange = $true
      } 
      if ($UpdatePath.ToUpper().EndsWith("/$channelName")) {
         $newUpdatePath = $newUpdatePath -replace "\/$channelName", "/$branchShortName"
         $newUpdateLong = $newUpdateLong -replace "\\$channelName\\", "\$branchName\"
         $madeChange = $true
      }
      if ($UpdatePath.ToUpper().Contains("/$channelName/")) {
         $newUpdatePath = $newUpdatePath -replace "\/$channelName\/", "/$branchShortName/"
         $newUpdateLong = $newUpdateLong -replace "\/$channelName\/", "/$branchName/"
         $madeChange = $true
      }
   }

   foreach ($channelName in $channelLongNames) {
      $channelName = $channelName.ToString().ToUpper()
      if ($UpdatePath.ToUpper().EndsWith("\$channelName")) {
         $newUpdatePath = $newUpdatePath -replace "\\$channelName", "\$branchShortName"
         $newUpdateLong = $newUpdateLong -replace "\\$channelName", "\$branchName"
         $madeChange = $true
      } 
      if ($UpdatePath.ToUpper().Contains("\$channelName\")) {
         $newUpdatePath = $newUpdatePath -replace "\\$channelName\\", "\$branchShortName\"
         $newUpdateLong = $newUpdateLong -replace "\\$channelName\\", "\$branchName\"
         $madeChange = $true
      } 
      if ($UpdatePath.ToUpper().EndsWith("/$channelName")) {
         $newUpdatePath = $newUpdatePath -replace "\/$channelName", "/$branchShortName"
         $newUpdateLong = $newUpdateLong -replace "\\$channelName\\", "\$branchName\"
         $madeChange = $true
      }
      if ($UpdatePath.ToUpper().Contains("/$channelName/")) {
         $newUpdatePath = $newUpdatePath -replace "\/$channelName\/", "/$branchShortName/"
         $newUpdateLong = $newUpdateLong -replace "\/$channelName\/", "/$branchName/"
         $madeChange = $true
      }
   }

   if (!($madeChange)) {
      if ($newUpdatePath.Contains("/")) {
         if ($newUpdatePath.EndsWith("/")) {
           $newUpdatePath += "$branchShortName"
         } else {
           $newUpdatePath += "/$branchShortName"
         }
      }
      if ($newUpdatePath.Contains("\")) {
         if ($newUpdatePath.EndsWith("\")) {
           $newUpdatePath += "$branchShortName"
         } else {
           $newUpdatePath += "\$branchShortName"
         }
      }
   }

   if (!($madeChange)) {
      if ($newUpdateLong.Contains("/")) {
         if ($newUpdateLong.EndsWith("/")) {
           $newUpdateLong += "$branchName"
         } else {
           $newUpdateLong += "/$branchName"
         }
      }
      if ($newUpdateLong.Contains("\")) {
         if ($newUpdateLong.EndsWith("\")) {
           $newUpdateLong += "$branchName"
         } else {
           $newUpdateLong += "\$branchName"
         }
      }
   }

   try {
     $pathAlive = Test-UpdateSource -UpdateSource $newUpdatePath -ValidateUpdateSourceFiles $ValidateUpdateSourceFiles -Bitness $Bitness
   } catch {
     $pathAlive = $false
   }

     if (!($pathAlive)) {
        try {
           $pathAlive = Test-UpdateSource -UpdateSource $newUpdateLong -ValidateUpdateSourceFiles $ValidateUpdateSourceFiles -Bitness $Bitness
        } catch {
            $pathAlive = $false
        }
        if ($pathAlive) {
           $newUpdatePath = $newUpdateLong
        }
     }
   
   if ($pathAlive) {
     return $newUpdatePath
   } else {
     return $UpdatePath
   }
}

Function Test-UpdateSource() {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [string] $UpdateSource = $NULL,

        [Parameter()]
        [bool] $ValidateUpdateSourceFiles = $true,

        [Parameter()]
        [string[]] $OfficeLanguages = $null,

        [Parameter()]
        [String] $Bitness = $NULL
    )

  	$uri = [System.Uri]$UpdateSource

    [bool]$sourceIsAlive = $false

    if($uri.Host){
	    $sourceIsAlive = Test-Connection -Count 1 -computername $uri.Host -Quiet
    }else{
        $sourceIsAlive = Test-Path $uri.LocalPath -ErrorAction SilentlyContinue
    }

    if ($ValidateUpdateSourceFiles) {
       if ($sourceIsAlive) {
           [string]$strIsAlive = Validate-UpdateSource -UpdateSource $UpdateSource -OfficeLanguages $OfficeLanguages -Bitness $Bitness
           if ($strIsAlive.ToLower() -eq "true") {
              $sourceIsAlive = $true
           } else {
              $sourceIsAlive = $false
           }
       }
    }

    return $sourceIsAlive
}

Function Validate-UpdateSource() {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [string] $UpdateSource = $NULL,

        [Parameter()]
        [string] $Bitness = $NULL,

        [Parameter()]
        [string[]] $OfficeLanguages = $NULL
    )

    [bool]$validUpdateSource = $true
    [string]$cabPath = ""

    if ($UpdateSource) {
        $mainRegPath = Get-OfficeCTRRegPath

        if(!$Bitness){
            $Bitness = "32"
        }

        $currentplatform = $Bitness

        if ($currentplatform -eq "x64") {
            $mainCab = "$UpdateSource\Office\Data\v64.cab"
            $Bitness = "64"
        }
        else{
            $mainCab = "$UpdateSource\Office\Data\v32.cab"
        }

        if ($mainRegPath) {
            $configRegPath = $mainRegPath + "\Configuration"
            $currentplatform = (Get-ItemProperty HKLM:\$configRegPath -Name Platform -ErrorAction SilentlyContinue).Platform
            $updateToVersion = (Get-ItemProperty HKLM:\$configRegPath -Name UpdateToVersion -ErrorAction SilentlyContinue).UpdateToVersion
            $llcc = (Get-ItemProperty HKLM:\$configRegPath -Name ClientCulture -ErrorAction SilentlyContinue).ClientCulture
        }

        if (!($updateToVersion)) {
           $cabXml = Get-CabVersion -FilePath $mainCab
           if ($cabXml) {
               $updateToVersion = $cabXml.Version.Available.Build
           }
        }

        [xml]$xml = Get-ChannelXml -Bitness $Bitness
        if ($OfficeLanguages) {
          $languages = $OfficeLanguages
        } else {
          $languages = Get-InstalledLanguages
        }

        $checkFiles = $xml.UpdateFiles.File | Where {   $_.language -eq "0" }
        foreach ($language in $languages) {
           $checkFiles += $xml.UpdateFiles.File | Where { $_.language -eq $language.LCID}
        }

        foreach ($checkFile in $checkFiles) {
           $fileName = $checkFile.name -replace "%version%", $updateToVersion
           $relativePath = $checkFile.relativePath -replace "%version%", $updateToVersion

           $fullPath = "$UpdateSource$relativePath$fileName"
           if ($fullPath.ToLower().StartsWith("http")) {
              $fullPath = $fullPath -replace "\\", "/"
           } else {
              $fullPath = $fullPath -replace "/", "\"
           }
           
           $updateFileExists = $false
           if ($fullPath.ToLower().StartsWith("http")) {
               $updateFileExists = Test-URL -url $fullPath
           } else {
               if ($fullPath.StartsWith("\\")) {
                  $updateFileExists = Test-ItemPathUNC -Path $fullPath
               } else {
                  $updateFileExists = Test-Path -Path $fullPath
               }
           }

           if (!($updateFileExists)) {
              $fileExists = $missingFiles.Contains($fullPath)
              if (!($fileExists)) {
                 $missingFiles.Add($fullPath)
                 Write-Host "Source File Missing: $fullPath"
                 Write-Log -Message "Source File Missing: $fullPath" -severity 1 -component "Office 365 Update Anywhere" 
              }     
              $validUpdateSource = $false
           }
        }

    }
    
    return $validUpdateSource
}

Function Copy-OfficeSourceFiles() {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [string] $Path = $NULL,

        [Parameter(Mandatory=$true)]
        [string] $Destination = $NULL,

        [Parameter()]
        [string] $Bitness = "x86",

        [Parameter()]
        [string[]] $OfficeLanguages = $null
    )

    [bool]$validUpdateSource = $true
    [string]$cabPath = ""

    if (($Path) -and ($Destination)) {
        $mainRegPath = Get-OfficeCTRRegPath
        if ($mainRegPath) {
            $configRegPath = $mainRegPath + "\Configuration"
            $currentplatform = (Get-ItemProperty HKLM:\$configRegPath -Name Platform -ErrorAction SilentlyContinue).Platform
            $updateToVersion = (Get-ItemProperty HKLM:\$configRegPath -Name UpdateToVersion -ErrorAction SilentlyContinue).UpdateToVersion
            $llcc = (Get-ItemProperty HKLM:\$configRegPath -Name ClientCulture -ErrorAction SilentlyContinue).ClientCulture
        }

        $currentplatform = $Bitness

        $mainCab = "$Path\Office\Data\v32.cab"
        $bitness = "32"
        if ($currentplatform -eq "x64") {
            $mainCab = "$Path\Office\Data\v64.cab"
            $bitness = "64"
        }

        if (!($updateToVersion)) {
           $cabXml = Get-CabVersion -FilePath $mainCab
           if ($cabXml) {
               $updateToVersion = $cabXml.Version.Available.Build
           }
        }

        [xml]$xml = Get-ChannelXml -Bitness $bitness
        if ($OfficeLanguages) {
          $languages = $OfficeLanguages
        } else {
          $languages = Get-InstalledLanguages
        }

        $checkFiles = $xml.UpdateFiles.File | Where {   $_.language -eq "0" }
        foreach ($language in $languages) {
           $checkFiles += $xml.UpdateFiles.File | Where { $_.language -eq $language.LCID}
        }

        foreach ($checkFile in $checkFiles) {
           $fileName = $checkFile.name -replace "%version%", $updateToVersion
           $relativePath = $checkFile.relativePath -replace "%version%", $updateToVersion

           $fullPath = "$UpdateSource$relativePath$fileName"
           if ($fullPath.ToLower().StartsWith("http")) {
              $fullPath = $fullPath -replace "\\", "/"
           } else {
              $fullPath = $fullPath -replace "/", "\"
           }
           
           $updateFileExists = $false
           if ($fullPath.ToLower().StartsWith("http")) {
               $updateFileExists = Test-URL -url $fullPath
           } else {
               if ($fullPath.StartsWith("\\")) {
                  $updateFileExists = Test-ItemPathUNC -Path $fullPath
               } else {
                  $updateFileExists = Test-Path -Path $fullPath
               }
           }

           if (!($updateFileExists)) {
              $fileExists = $missingFiles.Contains($fullPath)
              if (!($fileExists)) {
                 $missingFiles.Add($fullPath)
                 Write-Host "Source File Missing: $fullPath"
                 Write-Log -Message "Source File Missing: $fullPath" -severity 1 -component "Office 365 Update Anywhere" 
              }     
              $validUpdateSource = $false
           }
        }

    }
    
    return $validUpdateSource
}


function Detect-Channel {
   param( 

   )

   Process {
      $currentBaseUrl = Get-OfficeCDNUrl
      $channelXml = Get-ChannelXml

      $currentChannel = $channelXml.UpdateFiles.baseURL | Where {$_.URL -eq $currentBaseUrl -and $_.branch -notcontains 'Business' }
      return $currentChannel
   }

}

function Get-CabVersion {
   [CmdletBinding()]
   param( 
      [Parameter(Mandatory=$true)]
      [string] $FilePath = $NULL
   )

   process {
       $cabPath = $FilePath
       $fileName = Split-Path -Path $cabPath -Leaf
       $XMLFilePath = ""

       if ($cabPath.ToLower().StartsWith("http")) {
           $webclient = New-Object System.Net.WebClient
           $XMLFilePath = "$env:TEMP/$fileName"
           $XMLDownloadURL= $FilePath
           $webclient.DownloadFile($XMLDownloadURL,$XMLFilePath)
       } else {
         if ($cabPath.StartsWith("\\")) {
             if (Test-ItemPathUNC -Path $cabPath) {
                 $XMLFilePath = $cabPath
             }
         } else {
             if (Test-Path -Path $cabPath) {
                 $XMLFilePath = $cabPath
             }
         }
       }

       if ($XMLFilePath) {
           $tmpName = "VersionDescriptor.xml"
           expand $XMLFilePath $env:TEMP -f:$tmpName | Out-Null
           $tmpName = $env:TEMP + "\VersionDescriptor.xml"
           [xml]$versionXml = Get-Content $tmpName
           return $versionXml
       }
       return $null
   }
}

Function GetScriptRoot() {
 process {
     [string]$scriptPath = "."

     if ($PSScriptRoot) {
       $scriptPath = $PSScriptRoot
     } else {
       $scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
       $scriptPath = (Get-Item -Path ".\").FullName
     }

     return $scriptPath
 }
}

Function formatTimeItem() {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [string] $TimeItem = ""
    )

    [string]$returnItem = $TimeItem
    if ($TimeItem.Length -eq 1) {
       $returnItem = "0" + $TimeItem
    }
    return $returnItem
}

Function getOperationTime() {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [DateTime] $OperationStart
    )

    $operationTime = ""

    $dateDiff = NEW-TIMESPAN –Start $OperationStart –End (GET-DATE)
    $strHours = formatTimeItem -TimeItem $dateDiff.Hours.ToString() 
    $strMinutes = formatTimeItem -TimeItem $dateDiff.Minutes.ToString() 
    $strSeconds = formatTimeItem -TimeItem $dateDiff.Seconds.ToString() 

    if ($dateDiff.Days -gt 0) {
        $operationTime += "Days: " + $dateDiff.Days.ToString() + ":"  + $strHours + ":" + $strMinutes + ":" + $strSeconds
    }
    if ($dateDiff.Hours -gt 0 -and $dateDiff.Days -eq 0) {
        if ($operationTime.Length -gt 0) { $operationTime += " " }
        $operationTime += "Hours: " + $strHours + ":" + $strMinutes + ":" + $strSeconds
    }
    if ($dateDiff.Minutes -gt 0 -and $dateDiff.Days -eq 0 -and $dateDiff.Hours -eq 0) {
        if ($operationTime.Length -gt 0) { $operationTime += " " }
        $operationTime += "Minutes: " + $strMinutes + ":" + $strSeconds
    }
    if ($dateDiff.Seconds -gt 0 -and $dateDiff.Days -eq 0 -and $dateDiff.Hours -eq 0 -and $dateDiff.Minutes -eq 0) {
        if ($operationTime.Length -gt 0) { $operationTime += " " }
        $operationTime += "Seconds: " + $strSeconds
    }

    return $operationTime
}

Function Wait-ForOfficeCTRUpadate() {
    [CmdletBinding()]
    Param(
        [Parameter()]
        [int] $TimeOutInMinutes = 120
    )

    begin {
        $HKLM = [UInt32] "0x80000002"
        $HKCR = [UInt32] "0x80000000"
    }

    process {
       Write-Host "Waiting for Update process to Complete..."

       [datetime]$operationStart = Get-Date
       [datetime]$totalOperationStart = Get-Date

       Start-Sleep -Seconds 10

       $mainRegPath = Get-OfficeCTRRegPath
       $scenarioPath = $mainRegPath + "\scenario"

       $regProv = Get-Wmiobject -list "StdRegProv" -namespace root\default -ErrorAction Stop

       [DateTime]$startTime = Get-Date

       [string]$executingScenario = ""
       $failure = $false
       $cancelled = $false
       $updateRunning=$false
       [string[]]$trackProgress = @()
       [string[]]$trackComplete = @()
       [int]$noScenarioCount = 0

       do {
           $allComplete = $true
           $executingScenario = $regProv.GetStringValue($HKLM, $mainRegPath, "ExecutingScenario").sValue
           
           $scenarioKeys = $regProv.EnumKey($HKLM, $scenarioPath)
           foreach ($scenarioKey in $scenarioKeys.sNames) {
              if (!($executingScenario)) { continue }
              if ($scenarioKey.ToLower() -eq $executingScenario.ToLower()) {
                $taskKeyPath = Join-Path $scenarioPath "$scenarioKey\TasksState"
                $taskValues = $regProv.EnumValues($HKLM, $taskKeyPath).sNames

                foreach ($taskValue in $taskValues) {
                    [string]$status = $regProv.GetStringValue($HKLM, $taskKeyPath, $taskValue).sValue
                    $operation = $taskValue.Split(':')[0]
                    $keyValue = $taskValue
                   
                    if ($status.ToUpper() -eq "TASKSTATE_FAILED") {
                        $failure = $true
                    }

                    if ($status.ToUpper() -eq "TASKSTATE_CANCELLED") {
                        $cancelled = $true
                    }

                    if (($status.ToUpper() -eq "TASKSTATE_COMPLETED") -or`
                        ($status.ToUpper() -eq "TASKSTATE_CANCELLED") -or`
                        ($status.ToUpper() -eq "TASKSTATE_FAILED")) {
                        if (($trackProgress -contains $keyValue) -and !($trackComplete -contains $keyValue)) {
                            $displayValue = $operation + "`t" + $status + "`t" + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
                            #Write-Host $displayValue
                            $trackComplete += $keyValue 

                            $statusName = $status.Split('_')[1];

                            if (($operation.ToUpper().IndexOf("DOWNLOAD") -gt -1) -or `
                                ($operation.ToUpper().IndexOf("APPLY") -gt -1)) {

                                $operationTime = getOperationTime -OperationStart $operationStart

                                $displayText = $statusName + "`t" + $operationTime

                                Write-Host $displayText
                            }
                        }
                    } else {
                        $allComplete = $false
                        $updateRunning=$true


                        if (!($trackProgress -contains $keyValue)) {
                             $trackProgress += $keyValue 
                             $displayValue = $operation + "`t" + $status + "`t" + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')

                             $operationStart = Get-Date

                             if ($operation.ToUpper().IndexOf("DOWNLOAD") -gt -1) {
                                Write-Host "Downloading Update: " -NoNewline
                             }

                             if ($operation.ToUpper().IndexOf("APPLY") -gt -1) {
                                Write-Host "Applying Update: " -NoNewline
                             }

                             if ($operation.ToUpper().IndexOf("FINALIZE") -gt -1) {
                                Write-Host "Finalizing Update: " -NoNewline
                             }

                             #Write-Host $displayValue
                        }
                    }
                }
              }
           }

           if ($allComplete) {
              break;
           }

           if ($startTime -lt (Get-Date).AddHours(-$TimeOutInMinutes)) {
              throw "Waiting for Update Timed-Out"
              break;
           }

           Start-Sleep -Seconds 5
       } while($true -eq $true) 

       $operationTime = getOperationTime -OperationStart $operationStart

       $displayValue = ""
       if ($cancelled) {
         $displayValue = "CANCELLED`t" + $operationTime
       } else {
         if ($failure) {
            $displayValue = "FAILED`t" + $operationTime
         } else {
            $displayValue = "COMPLETED`t" + $operationTime
         }
       }

       Write-Host $displayValue

       $totalOperationTime = getOperationTime -OperationStart $totalOperationStart

       if ($updateRunning) {
          if ($failure) {
            Write-Host "Update Failed"
          } else {
            Write-Host "Update Completed - Total Time: $totalOperationTime"
          }
       } else {
          Write-Host "Update Not Running"
       } 
    }
}

Function Convert-Bool() {
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param
    (
        [Parameter(Mandatory=$true)]
        [bool] $value
    )

    $newValue = "$" + $value.ToString()
    return $newValue 
}

function Create-FileShare() {
    [CmdletBinding()]	
    Param
	(
		[Parameter()]
		[String]$Name = "",
		
		[Parameter()]
		[String]$Path = ""
	)

    $description = "$name"

    $Method = "Create"
    $sd = ([WMIClass] "Win32_SecurityDescriptor").CreateInstance()

    #AccessMasks:
    #2032127 = Full Control
    #1245631 = Change
    #1179817 = Read

    $userName = "$env:USERDOMAIN\$env:USERNAME"

    #Share with the user
    $ACE = ([WMIClass] "Win32_ACE").CreateInstance()
    $Trustee = ([WMIClass] "Win32_Trustee").CreateInstance()
    $Trustee.Name = $userName
    $Trustee.Domain = $NULL
    #original example assigned this, but I found it worked better if I left it empty
    #$Trustee.SID = ([wmi]"win32_userAccount.Domain='york.edu',Name='$name'").sid   
    $ace.AccessMask = 2032127
    $ace.AceFlags = 3 #Should almost always be three. Really. don't change it.
    $ace.AceType = 0 # 0 = allow, 1 = deny
    $ACE.Trustee = $Trustee 
    $sd.DACL += $ACE.psObject.baseobject 

    #Share with Domain Admins
    $ACE = ([WMIClass] "Win32_ACE").CreateInstance()
    $Trustee = ([WMIClass] "Win32_Trustee").CreateInstance()
    $Trustee.Name = "Domain Admins"
    $Trustee.Domain = $Null
    #$Trustee.SID = ([wmi]"win32_userAccount.Domain='york.edu',Name='$name'").sid   
    $ace.AccessMask = 2032127
    $ace.AceFlags = 3
    $ace.AceType = 0
    $ACE.Trustee = $Trustee 
    $sd.DACL += $ACE.psObject.baseobject    
    
     #Share with the user
    $ACE = ([WMIClass] "Win32_ACE").CreateInstance()
    $Trustee = ([WMIClass] "Win32_Trustee").CreateInstance()
    $Trustee.Name = "Everyone"
    $Trustee.Domain = $Null
    #original example assigned this, but I found it worked better if I left it empty
    #$Trustee.SID = ([wmi]"win32_userAccount.Domain='york.edu',Name='$name'").sid   
    $ace.AccessMask = 1179817 
    $ace.AceFlags = 3 #Should almost always be three. Really. don't change it.
    $ace.AceType = 0 # 0 = allow, 1 = deny
    $ACE.Trustee = $Trustee 
    $sd.DACL += $ACE.psObject.baseobject    

    $mc = [WmiClass]"Win32_Share"
    $InParams = $mc.psbase.GetMethodParameters($Method)
    $InParams.Access = $sd
    $InParams.Description = $description
    $InParams.MaximumAllowed = $Null
    $InParams.Name = $name
    $InParams.Password = $Null
    $InParams.Path = $path
    $InParams.Type = [uint32]0

    $R = $mc.PSBase.InvokeMethod($Method, $InParams, $Null)
    switch ($($R.ReturnValue))
     {
          0 { break}
          2 {Write-Host "Share:$name Path:$path Result:Access Denied" -foregroundcolor red -backgroundcolor yellow;break}
          8 {Write-Host "Share:$name Path:$path Result:Unknown Failure" -foregroundcolor red -backgroundcolor yellow;break}
          9 {Write-Host "Share:$name Path:$path Result:Invalid Name" -foregroundcolor red -backgroundcolor yellow;break}
          10 {Write-Host "Share:$name Path:$path Result:Invalid Level" -foregroundcolor red -backgroundcolor yellow;break}
          21 {Write-Host "Share:$name Path:$path Result:Invalid Parameter" -foregroundcolor red -backgroundcolor yellow;break}
          22 {Write-Host "Share:$name Path:$path Result:Duplicate Share" -foregroundcolor red -backgroundcolor yellow;break}
          23 {Write-Host "Share:$name Path:$path Result:Reedirected Path" -foregroundcolor red -backgroundcolor yellow;break}
          24 {Write-Host "Share:$name Path:$path Result:Unknown Device or Directory" -foregroundcolor red -backgroundcolor yellow;break}
          25 {Write-Host "Share:$name Path:$path Result:Network Name Not Found" -foregroundcolor red -backgroundcolor yellow;break}
          default {Write-Host "Share:$name Path:$path Result:*** Unknown Error ***" -foregroundcolor red -backgroundcolor yellow;break}
     }
}

function Get-Fileshare() {
    [CmdletBinding()]	
    Param
	(
		[Parameter()]
		[String]$Name = ""
	)

    $share = Get-WmiObject Win32_Share | where { $_.Name -eq $Name }

    if ($share) {
        return $share;
    }

    return $null
}

function Check-AdminAccess() {
If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`    [Security.Principal.WindowsBuiltInRole] “Administrator”)){    throw “You do not have Administrator rights to run this script!`nPlease re-run this script as an Administrator!”}
}

function Get-LargestDrive() {
   [CmdletBinding()]
   param( 
   )
   process {
      $drives = Get-Partition | where {$_.DriveLetter}
      $driveInfoList = @()

      foreach ($drive in $drives) {
          $driveLetter = $drive.DriveLetter
          $deviceFilter = "DeviceID='" + $driveLetter + ":'" 
 
          $driveInfo = Get-WmiObject Win32_LogicalDisk -ComputerName "." -Filter $deviceFilter
          $driveInfoList += $driveInfo
      }

      $SortList = Sort-Object -InputObject $driveInfoList -Property FreeSpace

      $FreeSpaceDrive = $SortList[0]
      return $FreeSpaceDrive.DeviceID
   }
}

function Get-ChannelXml() {
    [CmdletBinding()]	
    Param
	(
	    [Parameter()]
	    [string]$FolderPath = $null,

	    [Parameter()]
	    [bool]$OverWrite = $false,

        [Parameter()]
        [string] $Bitness = "32"
	)

   process {
       $cabPath = "$PSScriptRoot\ofl.cab"
       [bool]$downloadFile = $true

       if (!($OverWrite)) {
          if ($FolderPath) {
              $XMLFilePath = "$FolderPath\ofl.cab"
              if (Test-Path -Path $XMLFilePath) {
                 $downloadFile = $false
              } else {
                throw "File missing $FolderPath\ofl.cab"
              }
          }
       }

       if ($downloadFile) {
           $webclient = New-Object System.Net.WebClient
           $XMLFilePath = "$env:TEMP/ofl.cab"
           $XMLDownloadURL = "http://officecdn.microsoft.com/pr/wsus/ofl.cab"
           $webclient.DownloadFile($XMLDownloadURL,$XMLFilePath)

           if ($FolderPath) {
             [System.IO.Directory]::CreateDirectory($FolderPath) | Out-Null
             $targetFile = "$FolderPath\ofl.cab"
             Copy-Item -Path $XMLFilePath -Destination $targetFile -Force
           }
       }

       $tmpName = "o365client_" + $Bitness + "bit.xml"
       expand $XMLFilePath $env:TEMP -f:$tmpName | Out-Null
       $tmpName = $env:TEMP + "\" + $tmpName
       
       [xml]$channelXml = Get-Content $tmpName

       return $channelXml
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

Function Get-LatestVersion() {
  [CmdletBinding()]
  Param(
     [Parameter(Mandatory=$true)]
     [string] $UpdateURLPath
  )

  process {
    [array]$totalVersion = @()
    $Version = $null

    $LatestBranchVersionPath = $UpdateURLPath + '\Office\Data'
    if(Test-Path $LatestBranchVersionPath){
        $DirectoryList = Get-ChildItem $LatestBranchVersionPath
        Foreach($listItem in $DirectoryList){
            if($listItem.GetType().Name -eq 'DirectoryInfo'){
                $totalVersion+=$listItem.Name
            }
        }
    }

    $totalVersion = $totalVersion | Sort-Object -Descending
    
    #sets version number to the newest version in directory for channel if version is not set by user in argument  
    if($totalVersion.Count -gt 0){
        $Version = $totalVersion[0]
    }

    return $Version
  }
}

function Get-ChannelLatestVersion() {
   [CmdletBinding()]
   param( 
      [Parameter(Mandatory=$true)]
      [string]$ChannelUrl,

      [Parameter(Mandatory=$true)]
      [string]$Channel,

	  [Parameter()]
	  [string]$FolderPath = $null,

	  [Parameter()]
	  [bool]$OverWrite = $false
   )

   process {

       [bool]$downloadFile = $true

       $channelShortName = ConvertChannelNameToShortName -ChannelName $Channel

       if (!($OverWrite)) {
          if ($FolderPath) {
              $CABFilePath = "$FolderPath\$channelShortName\v32.cab"

              if (!(Test-Path -Path $CABFilePath)) {
                 $CABFilePath = "$FolderPath\$channelShortName\v64.cab"
              }

              if (Test-Path -Path $CABFilePath) {
                 $downloadFile = $false
              } else {
                throw "File missing $FolderPath\$channelShortName\v64.cab or $FolderPath\$channelShortName\v64.cab"
              }
          }
       }

       if ($downloadFile) {
           $webclient = New-Object System.Net.WebClient
           $CABFilePath = "$env:TEMP/v32.cab"
           $XMLDownloadURL = "$ChannelUrl/Office/Data/v32.cab"
           $webclient.DownloadFile($XMLDownloadURL,$CABFilePath)

           if ($FolderPath) {
             [System.IO.Directory]::CreateDirectory($FolderPath) | Out-Null

             $channelShortName = ConvertChannelNameToShortName -ChannelName $Channel 

             $targetFile = "$FolderPath\$channelShortName\v32.cab"
             Copy-Item -Path $CABFilePath -Destination $targetFile -Force
           }
       }

       $tmpName = "VersionDescriptor.xml"
       expand $CABFilePath $env:TEMP -f:$tmpName | Out-Null
       $tmpName = $env:TEMP + "\VersionDescriptor.xml"
       [xml]$versionXml = Get-Content $tmpName

       return $versionXml.Version.Available.Build
   }
}

function ConvertChannelNameToShortName {
    Param(
       [Parameter()]
       [string] $ChannelName
    )
    Process {
       if ($ChannelName.ToLower() -eq "FirstReleaseCurrent".ToLower()) {
         return "FRCC"
       }
       if ($ChannelName.ToLower() -eq "Current".ToLower()) {
         return "CC"
       }
       if ($ChannelName.ToLower() -eq "FirstReleaseDeferred".ToLower()) {
         return "FRDC"
       }
       if ($ChannelName.ToLower() -eq "Deferred".ToLower()) {
         return "DC"
       }
       if ($ChannelName.ToLower() -eq "Business".ToLower()) {
         return "DC"
       }
       if ($ChannelName.ToLower() -eq "FirstReleaseBusiness".ToLower()) {
         return "FRDC"
       }
    }
}

function Check-FileDependencies() {
   [CmdletBinding()]
   param( 
      [Parameter(Mandatory=$true)]
      [string[]]$Files
   )

   process {
      foreach ($file in $Files) {
        $fileExists = Test-ItemPathUNC -Path $file
        if (!($fileExists)) {
                throw "Missing Dependency File $file"    
        }
        . $file
      }
   }
}

function ImportDeploymentDependencies() {
   [CmdletBinding()]
   param( 
      [Parameter(Mandatory=$true)]
      [string]$ScriptPath
   )
   process {
       #Importing all required functions
       $dependFiles = @(  "$scriptPath\Generate-ODTConfigurationXML.ps1"
                          "$scriptPath\Edit-OfficeConfigurationFile.ps1"
                          "$scriptPath\Install-OfficeClickToRun.ps1"
                          "$scriptPath\SharedFunctions.ps1"
                          )

       foreach ($dependFile in $dependFiles) {
          Check-FileDependencies -Files $dependFiles
       }
   }
}

function UpdateConfigurtionXml() {
   [CmdletBinding()]
   param( 
      [Parameter(Mandatory=$true)]
      [string[]]$Files
   )
   process {
     $languages = Get-XMLLanguages -Path $targetFilePath

     if (Test-UpdateSource -UpdateSource $UpdateURLPath -OfficeLanguages $languages) {
         Set-ODTAdd -TargetFilePath $targetFilePath -SourcePath $UpdateURLPath | Out-Null
     }

     if (($Bitness -eq "32") -or ($Bitness -eq "x86")) {
         Set-ODTAdd -TargetFilePath $targetFilePath -Bitness 32 | Out-Null
     } else {
         Set-ODTAdd -TargetFilePath $targetFilePath -Bitness 64 | Out-Null
     }
   }
}

function Locate-UpdateSource() {
   [CmdletBinding()]
   param( 
      [Parameter(Mandatory=$true)]
      [string]$UpdateURLPath,

      [Parameter(Mandatory=$true)]
      [string]$SourceFileFolder,

      [Parameter(Mandatory=$true)]
      [string] $Channel = $null,

      [Parameter(Mandatory=$false)]
      [string] $Bitness = $null
   )
   process {
     if ($SourceFileFolder) {
       if (Test-ItemPathUNC -Path "$UpdateURLPath\$SourceFileFolder") {
          $UpdateURLPath = "$UpdateURLPath\$SourceFileFolder"
       }
     }

     $UpdateURLPath = Change-UpdatePathToChannel -Channel $Channel -UpdatePath $UpdateURLPath -Bitness $Bitness
     return $UpdateURLPath
   }
}

function Update-ConfigurationXml() {
   [CmdletBinding()]
   param(
      [Parameter(Mandatory=$true)]
      [string] $TargetFilePath,

      [Parameter(Mandatory=$true)]
      [string] $UpdateURLPath
   )
   process {
      $scriptPath = GetScriptRoot
      $editFilePath = "$scriptPath\Edit-OfficeConfigurationFile.ps1"

      $languages = Get-XMLLanguages -Path $TargetFilePath

      if (Test-Path -Path $editFilePath) {
          . $editFilePath

          if (Test-Path -Path "$UpdateURLPath\Office\Data") {
              if (Test-UpdateSource -UpdateSource $UpdateURLPath -OfficeLanguages $languages) {
                 Set-ODTAdd -TargetFilePath $TargetFilePath -SourcePath $UpdateURLPath | Out-Null
              }
          }

          if (($Bitness -eq "32") -or ($Bitness -eq "x86")) {
             Set-ODTAdd -TargetFilePath $TargetFilePath -Bitness 32 | Out-Null
          } else {
             Set-ODTAdd -TargetFilePath $TargetFilePath -Bitness 64 | Out-Null
          }
      }
   }
 }

function Exclude-Applications() {
   [CmdletBinding()]
   param(
      [Parameter(Mandatory=$true)]
      [string] $TargetFilePath,

      [Parameter(Mandatory=$true)]
      [string[]] $ExcludeApps
   )
   process {
      if ((Get-ODTProductToAdd -TargetFilePath $targetFilePath -ProductId O365ProPlusRetail)) {           Set-ODTProductToAdd -ProductId "O365ProPlusRetail" -TargetFilePath $targetFilePath -ExcludeApps $ExcludeApps | Out-Null      }      if ((Get-ODTProductToAdd -TargetFilePath $targetFilePath -ProductId O365BusinessRetail)) {           Set-ODTProductToAdd -ProductId "O365BusinessRetail" -TargetFilePath $targetFilePath -ExcludeApps $ExcludeApps | Out-Null      }
   }
}

function Add-ProductSku() {
   [CmdletBinding()]
   param(
      [Parameter(Mandatory=$true)]
      [string] $TargetFilePath,

      [Parameter(Mandatory=$true)]
      [Microsoft.Office.Products[]] $ProductIDs,

      [Parameter(Mandatory=$true)]
      [string[]] $Languages
   )
   process {
     $scriptPath = GetScriptRoot
     $editFilePath = "$scriptPath\Edit-OfficeConfigurationFile.ps1"
     if (Test-Path -Path $editFilePath) {
          . $editFilePath
     }

    foreach ($ProductID in $ProductIDs) {        if (!(Get-ODTProductToAdd -TargetFilePath $targetFilePath -ProductId $ProductID)) {              Add-ODTProductToAdd -ProductId $ProductID -TargetFilePath $targetFilePath -LanguageIds $languages | Out-Null            }
    }
   }
}

function Remove-ProductSku() {
   [CmdletBinding()]
   param(
      [Parameter(Mandatory=$true)]
      [string] $TargetFilePath,

      [Parameter(Mandatory=$true)]
      [Microsoft.Office.Products[]] $ProductIDs
   )
   process {
     $scriptPath = GetScriptRoot
     $editFilePath = "$scriptPath\Edit-OfficeConfigurationFile.ps1"
     if (Test-Path -Path $editFilePath) {
          . $editFilePath
     }
    foreach ($ProductID in $ProductIDs) {        if (Get-ODTProductToAdd -TargetFilePath $targetFilePath -ProductId $ProductID) {            Remove-ODTProductToAdd -ProductId $ProductID -TargetFilePath $targetFilePath -LanguageIds $languages | Out-Null            }
    }
   }
}

function Add-ProductLanguage() {
   [CmdletBinding()]
   param(
      [Parameter(Mandatory=$true)]
      [string] $TargetFilePath,

      [Parameter(Mandatory=$true)]
      [Microsoft.Office.ProductSelection[]] $ProductIDs,

      [Parameter(Mandatory=$true)]
      [string[]] $Languages
   )
   process {
     $scriptPath = GetScriptRoot
     $editFilePath = "$scriptPath\Edit-OfficeConfigurationFile.ps1"
     if (Test-Path -Path $editFilePath) {
          . $editFilePath
     }

    if ($ProductIDs -eq "All") {
        $productsToCheck = @("O365ProPlusRetail","O365BusinessRetail","VisioProRetail","ProjectProRetail","SPDRetail","VisioProXVolume","VisioStdXVolume","ProjectProXVolume","ProjectStdXVolume")
         
        foreach ($ProductID in $productsToCheck) {            $existingSku = Get-ODTProductToAdd -TargetFilePath $targetFilePath -ProductId $ProductID            if ($existingSku) {                $newLangList = @()                foreach ($language in $existingSku.Languages) {                   $newLangList += $language                }                foreach ($newLanguage in $languages) {                   if (!($newLangList.Contains($newLanguage))) {                     $newLangList += $newLanguage                   }                }                if (Get-ODTProductToAdd -TargetFilePath $targetFilePath -ProductId $ProductID) {                   Set-ODTProductToAdd -ProductId $ProductID -TargetFilePath $targetFilePath -LanguageIds $newLangList | Out-Null                  }
            }
        }
    } else {
        foreach ($ProductID in $ProductIDs) {            if (!(Get-ODTProductToAdd -TargetFilePath $targetFilePath -ProductId $ProductID.ToString())) {                  Add-ODTProductToAdd -ProductId $ProductID.ToString() -TargetFilePath $targetFilePath -LanguageIds $languages | Out-Null                }

            $existingSku = Get-ODTProductToAdd -TargetFilePath $targetFilePath -ProductId $ProductID.ToString()            if ($existingSku) {                $newLangList = @()                foreach ($language in $existingSku.Languages) {                   $newLangList += $language                }                foreach ($newLanguage in $languages) {                   if (!($newLangList.Contains($newLanguage))) {                     $newLangList += $newLanguage                   }                }                if (Get-ODTProductToAdd -TargetFilePath $targetFilePath -ProductId $ProductID.ToString()) {                   Set-ODTProductToAdd -ProductId $ProductID.ToString() -TargetFilePath $targetFilePath -LanguageIds $newLangList | Out-Null                  }
            }
        }
    }
   }
}

function Remove-ProductLanguage() {
   [CmdletBinding()]
   param(
      [Parameter(Mandatory=$true)]
      [string] $TargetFilePath,

      [Parameter(Mandatory=$true)]
      [Microsoft.Office.ProductSelection[]] $ProductIDs,

      [Parameter(Mandatory=$true)]
      [string[]] $Languages
   )
   process {
     $scriptPath = GetScriptRoot
     $editFilePath = "$scriptPath\Edit-OfficeConfigurationFile.ps1"
     if (Test-Path -Path $editFilePath) {
          . $editFilePath
     }

    if ($ProductIDs -eq "All") {
        $productsToCheck = @("O365ProPlusRetail","O365BusinessRetail","VisioProRetail","ProjectProRetail","SPDRetail","VisioProXVolume","VisioStdXVolume","ProjectProXVolume","ProjectStdXVolume")
         
        foreach ($ProductID in $productsToCheck) {            $existingSku = Get-ODTProductToAdd -TargetFilePath $targetFilePath -ProductId $ProductID            if ($existingSku) {                $newLangList = @()                foreach ($language in $existingSku.Languages) {                   [bool]$addLanguages = $true                   foreach ($newLanguage in $languages) {                      if ($language.ToLower() -eq $newLanguage.ToLower()) {                         $addLanguages = $false                      }                   }                   if ($addLanguages) {                       $newLangList += $language                   }                }                if (Get-ODTProductToAdd -TargetFilePath $targetFilePath -ProductId $ProductID) {                   Set-ODTProductToAdd -ProductId $ProductID -TargetFilePath $targetFilePath -LanguageIds $newLangList | Out-Null                  }
            }
        }
    } else {
        foreach ($ProductID in $ProductIDs) {            if (!(Get-ODTProductToAdd -TargetFilePath $targetFilePath -ProductId $ProductID.ToString())) {                  Add-ODTProductToAdd -ProductId $ProductID.ToString() -TargetFilePath $targetFilePath -LanguageIds $languages | Out-Null                }

            $existingSku = Get-ODTProductToAdd -TargetFilePath $targetFilePath -ProductId $ProductID.ToString()            if ($existingSku) {                $newLangList = @()                foreach ($language in $existingSku.Languages) {                   [bool]$addLanguages = $true                   foreach ($newLanguage in $languages) {                      if ($language.ToLower() -eq $newLanguage.ToLower()) {                         $addLanguages = $false                      }                   }                   if ($addLanguages) {                       $newLangList += $language                   }                }                if (Get-ODTProductToAdd -TargetFilePath $targetFilePath -ProductId $ProductID.ToString()) {                   Set-ODTProductToAdd -ProductId $ProductID.ToString() -TargetFilePath $targetFilePath -LanguageIds $newLangList | Out-Null                  }
            }
        }
    }
   }
}
