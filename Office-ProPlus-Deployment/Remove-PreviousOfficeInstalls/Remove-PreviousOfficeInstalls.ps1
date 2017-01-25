[CmdletBinding(SupportsShouldProcess=$true)]
param(
[Parameter(ValueFromPipelineByPropertyName=$true)]
[bool]$RemoveClickToRunVersions = $false,

[Parameter(ValueFromPipelineByPropertyName=$true)]
[bool]$Remove2016Installs = $false,

[Parameter(ValueFromPipelineByPropertyName=$true)]
[bool]$Force = $false,

[Parameter(ValueFromPipelineByPropertyName=$true)]
[bool]$KeepUserSettings = $true,

[Parameter(ValueFromPipelineByPropertyName=$true)]
[bool]$KeepLync = $false,

[Parameter(ValueFromPipelineByPropertyName=$true)]
[bool]$NoReboot = $false
)

Function Remove-PreviousOfficeInstalls{
  [CmdletBinding(SupportsShouldProcess=$true)]
  param(
    [Parameter(ValueFromPipelineByPropertyName=$true)]
    [bool]$RemoveClickToRunVersions = $true,

    [Parameter(ValueFromPipelineByPropertyName=$true)]
    [bool]$Remove2016Installs = $false,

    [Parameter(ValueFromPipelineByPropertyName=$true)]
    [bool]$Force = $false,

    [Parameter(ValueFromPipelineByPropertyName=$true)]
    [bool]$KeepUserSettings = $true,

    [Parameter(ValueFromPipelineByPropertyName=$true)]
    [bool]$KeepLync = $false,

    [Parameter(ValueFromPipelineByPropertyName=$true)]
    [bool]$NoReboot = $false,

    [Parameter(ValueFromPipelineByPropertyName=$true)]
    [bool]$Quiet = $true
  )

  Process {
    $c2rVBS = "OffScrubc2r.vbs"
    $03VBS = "OffScrub03.vbs"
    $07VBS = "OffScrub07.vbs"
    $10VBS = "OffScrub10.vbs"
    $15MSIVBS = "OffScrub_O15msi.vbs"
    $16MSIVBS = "OffScrub_O16msi.vbs"

    if ($Quiet) {
      $argList = "CLIENTALL /QUIET"
    } else {
      $argList = "CLIENTALL"
    }
    
    if ($Force) {
        $argList += " /FORCE"
    }

    if ($KeepUserSettings) {
       $argList += " /KEEPUSERSETTINGS"
    } else {
       $argList += " /DELETEUSERSETTINGS"
    }

    if ($KeepLync) {
       $argList += " /KEEPLYNC"
    } else {
       $argList += " /REMOVELYNC"
    }

    if ($NoReboot) {
        $argList += " /NOREBOOT"
    }

    $scriptPath = GetScriptRoot

    Write-Host "Detecting Office installs..."
            #write log
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Detecting Office installs..."

    $officeVersions = Get-OfficeVersion -ShowAllInstalledProducts | select *
    $ActionFiles = @()
    
    $removeOffice = $true
    if (!( $officeVersions)) {
       Write-Host "Microsoft Office is not installed"
            #write log
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Microsoft Office is not installed"
       $removeOffice = $false
    }

    if ($removeOffice) {
        $osVersion = (Get-WmiObject -Class Win32_OperatingSystem).Version
        [int]$osVersion = $osVersion.Split('.')[0]
        if($osVersion -ge '10') {
            Remove-PinnedOfficeApplications
        }

        [bool]$office03Removed = $false
        [bool]$office07Removed = $false
        [bool]$office10Removed = $false
        [bool]$office15Removed = $false
        [bool]$office16Removed = $false
        [bool]$officeC2RRemoved = $false

        [bool]$c2r2013Installed = $false
        [bool]$c2r2016Installed = $false

        foreach ($officeVersion in $officeVersions) {
           if($officeVersion.ClicktoRun.ToLower() -eq "true"){
              if ($officeVersion.Version -like '15.*') {
                  $c2r2013Installed = $true
              }
              if ($officeVersion.Version -like '16.*') {
                  $c2r2016Installed = $true
              }
           }
        }

        foreach ($officeVersion in $officeVersions) {
            if($officeVersion.ClicktoRun.ToLower() -eq "true"){
              $removeC2R = $false

              if (!($officeC2RRemoved)) {
                  if ($RemoveClickToRunVersions -and (!($c2r2016Installed))) {
                     $removeC2R = $true
                  }
                  if ($Remove2016Installs -and $RemoveClickToRunVersions) {
                     $removeC2R = $true
                  }
              }

              if ($removeC2R) {
                  Write-Host "`tRemoving Office Click-To-Run..."
            #write log
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Removing Office Click-To-Run..."
                  $ActionFile = "$scriptPath\$c2rVBS"
                  $cmdLine = """$ActionFile"" $argList"
                 
                  if (Test-Path -Path $ActionFile) {
                    $cmd = "cmd /c cscript //Nologo $cmdLine"
                    Invoke-Expression $cmd
                    $officeC2RRemoved = $true
                    $c2r2013Installed = $false
                  } else {
            #write log
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Required file missing: $ActionFile"
                    throw "Required file missing: $ActionFile"
                  }
                  Write-Host ""
              }

            }
        }

        foreach ($officeVersion in $officeVersions) {
            if($officeVersion.ClicktoRun.ToLower() -ne "true"){
                #Set script file based on office version, if no office detected continue to next computer skipping this one.
                switch -wildcard ($officeVersion.Version)
                {
                    "11.*"
                    {
                        if (!($office03Removed)) {
                            Write-Host "`tRemoving Office 2003..."
            #write log
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Removing Office 2003..."
                            $ActionFile = "$scriptPath\$03VBS"
                            $cmdLine = """$ActionFile"" $argList"
                        
                            if (Test-Path -Path $ActionFile) {
                                $cmd = "cmd /c cscript //Nologo $cmdLine"
                                Invoke-Expression $cmd
                                $office03Removed = $true
                            } else {
            #write log
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Required file missing: $ActionFile"
                               throw "Required file missing: $ActionFile"
                            }
                            Write-Host ""
                        }
                    }
                    "12.*"
                    {
                        if (!($office07Removed)) {
                            Write-Host "`tRemoving Office 2007..."
            #write log
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Removing Office 2007..."
                            $ActionFile = "$scriptPath\$07VBS"
                            $cmdLine = """$ActionFile"" $argList"
                        
                            if (Test-Path -Path $ActionFile) {
                                $cmd = "cmd /c cscript //Nologo $cmdLine"
                                Invoke-Expression $cmd
                                $office07Removed = $true
                            } else {
            #write log
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Required file missing $ActionFile"
                               throw "Required file missing: $ActionFile"
                            }
                            Write-Host ""
                        }
                    }
                    "14.*"
                    {
                        if (!($office10Removed)) {
                            Write-Host "`tRemoving Office 2010..."
            #write log
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Removing Office 2010..."
                            $ActionFile = "$scriptPath\$10VBS"
                            $cmdLine = """$ActionFile"" $argList"
                        
                            if (Test-Path -Path $ActionFile) {
                                $cmd = "cmd /c cscript //Nologo $cmdLine"
                                Invoke-Expression $cmd
                                $office10Removed = $true
                            } else {
            #write log
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Required file missing: $ActionFile"
                               throw "Required file missing: $ActionFile"
                            }
                            Write-Host ""
                        }
                    }
                    "15.*"
                    {
                        if (!($office15Removed)) {
                            if (!($c2r2013Installed)) {
                                Write-Host "`tRemoving Office 2013..."
            #write log
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Removing Office 2013..."
                                $ActionFile = "$scriptPath\$15MSIVBS"
                                $cmdLine = """$ActionFile"" $argList"
                        
                                if (Test-Path -Path $ActionFile) {
                                   $cmd = "cmd /c cscript //Nologo $cmdLine"
                                   Invoke-Expression $cmd 
                                   $office15Removed = $true
                                } else {
            #write log
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Required file missing: $ActionFile"
                                   throw "Required file missing: $ActionFile"
                                }
                                Write-Host ""
                            } else {
            #write log
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Office 2013 cannot be removed if 2013 Click-To-Run is installed. Use the -RemoveClickToRunVersions parameter to remove Click-To-Run installs."
                              throw "Office 2013 cannot be removed if 2013 Click-To-Run is installed. Use the -RemoveClickToRunVersions parameter to remove Click-To-Run installs."
                            }
                        }
                    }
                    "16.*"
                    {
                       if (!($office16Removed)) {
                           if ($Remove2016Installs) {

                                if (!($c2r2016Installed)) {
                                      Write-Host "`tRemoving Office 2016..."
            #write log
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Removing Office 2016..."
                                      $ActionFile = "$scriptPath\$16MSIVBS"
                                      $cmdLine = """$ActionFile"" $argList"
                          
                                      if (Test-Path -Path $ActionFile) {
                                        $cmd = "cmd /c cscript //Nologo $cmdLine"
                                        Invoke-Expression $cmd
                                        $office16Removed = $true
                                      } else {
            #write log
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Required file missing: $ActionFile"
                                        throw "Required file missing: $ActionFile"
                                      }
                                      Write-Host ""
                                } else {
            #write log
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Office 2016 cannot be removed if 2016 Click-To-Run is installed. Use the -RemoveClickToRunVersions parameter to remove Click-To-Run installs."
                                  throw "Office 2016 cannot be removed if 2016 Click-To-Run is installed. Use the -RemoveClickToRunVersions parameter to remove Click-To-Run installs."
                                }

                           }
                       }
                    }
                    default 
                    {
                        continue
                    }
                }
            }
        }
    }
  }
}

Function IsDotSourced() {
  [CmdletBinding(SupportsShouldProcess=$true)]
  param(
    [Parameter(ValueFromPipelineByPropertyName=$true)]
    [string]$InvocationLine = ""
  )
  $cmdLine = $InvocationLine.Trim()
  Do {
    $cmdLine = $cmdLine.Replace(" ", "")
  } while($cmdLine.Contains(" "))

  $dotSourced = $false
  if ($cmdLine -match '^\.\\') {
     $dotSourced = $false
  } else {
     $dotSourced = ($cmdLine -match '^\.')
  }

  return $dotSourced
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
DateUpdated: 2016-07-20
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

           if ($ConfigItemList.Contains($key.ToUpper()) -and $name.ToUpper().Contains("MICROSOFT OFFICE") `
                                                        -and $name.ToUpper() -notlike "*MUI*" `
                                                        -and $name.ToUpper() -notlike "*VISIO*" `
                                                        -and $name.ToUpper() -notlike "*PROJECT*" `
                                                        -and $name.ToUpper() -notlike "*PROOFING*") {
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
        $fileName = $_.InvocationInfo.ScriptName.Substring($_.InvocationInfo.ScriptName.LastIndexOf("\")+1)
        WriteToLogFile -LNumber $_.InvocationInfo.ScriptLineNumber -FName $fileName -ActionError $_
    }
}

$dotSourced = IsDotSourced -InvocationLine $MyInvocation.Line

if (!($dotSourced)) {
   Remove-PreviousOfficeInstalls -RemoveClickToRunVersions $RemoveClickToRunVersions -Remove2016Installs $Remove2016Installs -Force $Force -KeepUserSettings $KeepUserSettings -KeepLync $KeepLync -NoReboot $NoReboot
}