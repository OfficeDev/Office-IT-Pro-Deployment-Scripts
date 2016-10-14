[CmdletBinding()]
Param(
    [Parameter()]
    [bool] $WaitForUpdateToFinish = $true,

    [Parameter()]
    [bool] $EnableUpdateAnywhere = $true,

    [Parameter()]
    [bool] $ForceAppShutdown = $true,
    
    [Parameter()]
    [bool] $UpdatePromptUser = $false,

    [Parameter()]
    [bool] $DisplayLevel = $false,

    [Parameter()]
    [string] $UpdateToVersion = $NULL,

    [Parameter()]
    [string] $LogPath = $null,

    [Parameter()]
    [string] $LogName = $null,

    [Parameter()]
    [bool] $ValidateUpdateSourceFiles = $true,

    [Parameter()]
    [bool] $UseScriptLocationAsUpdateSource = $false

)

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

Function Update-Office365Anywhere() {
<#
.Synopsis
This function is designed to provide way for Office Click-To-Run clients to have the ability to update themselves from a managed network source
or from the Internet depending on the availability of the primary update source.

.DESCRIPTION
This function is designed to provide way for Office Click-To-Run clients to have the ability to update themselves from a managed network source
or from the Internet depending on the availability of the primary update source.  The idea behind this is if users have laptops and are mobile 
they may not recieve updates if they are not able to be in the office on regular basis.  This functionality is available with this function but it's 
use can be controller by the parameter -EnableUpdateAnywhere.  This function also provides a way to initiate an update and the script will wait
for the update to complete before exiting. Natively starting an update executable does not wait for the process to complete before exiting and
in certain scenarios it may be useful to have the update process wait for the update to complete.

.NOTES   
Name: Update-Office365Anywhere
Version: 1.1.0
DateCreated: 2015-08-28
DateUpdated: 2015-09-03

.LINK
https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts

.PARAMETER WaitForUpdateToFinish
If this parameter is set to $true then the function will monitor the Office update and will not exit until the update process has stopped.
If this parameter is set to $false then the script will exit right after the update process has been started.  By default this parameter is set
to $true

.PARAMETER EnableUpdateAnywhere
This parameter controls whether the UpdateAnywhere functionality is used or not. When enabled the update process will check the availbility
of the update source set for the client.  If that update source is not available then it will update the client from the Microsoft Office CDN.
When set to $false the function will only use the Update source configured on the client. By default it is set to $true.

.PARAMETER ForceAppShutdown
This specifies whether the user will be given the option to cancel out of the update. However, if this variable is set to True, then the applications will be shut down immediately and the update will proceed.

.PARAMETER UpdatePromptUser
This specifies whether or not the user will see this dialog before automatically applying the updates:

.PARAMETER DisplayLevel
This specifies whether the user will see a user interface during the update. Setting this to false will hide all update UI (including error UI that is encountered during the update scenario).

.PARAMETER UpdateToVersion
This specifies the version to which Office needs to be updated to.  This can used to install a newer or an older version than what is presently installed.

.PARAMETER ValidateUpdateSourceFiles
If this parameter is set to true then the script will ensure the update source has all the files necessary to perform the update

.EXAMPLE
Update-Office365Anywhere 

Description:
Will generate the Office Deployment Tool (ODT) configuration XML based on the local computer

#>

    [CmdletBinding()]
    Param(
        [Parameter()]
        [bool] $WaitForUpdateToFinish = $true,

        [Parameter()]
        [bool] $EnableUpdateAnywhere = $true,

        [Parameter()]
        [bool] $ForceAppShutdown = $false,

        [Parameter()]
        [bool] $UpdatePromptUser = $false,

        [Parameter()]
        [bool] $DisplayLevel = $false,

        [Parameter()]
        [string] $UpdateToVersion = $NULL,

        [Parameter()]
        [string] $LogPath = $NULL,

        [Parameter()]
        [string] $LogName = $NULL,
        
        [Parameter()]
        [bool] $ValidateUpdateSourceFiles = $true,

        [Parameter()]
        [bool] $UseScriptLocationAsUpdateSource = $false
    )

    Process {
        try {
            $Global:UpdateAnywhereLogPath = $LogPath;
            $Global:UpdateAnywhereLogFileName = $LogName;

            $scriptPath = GetScriptRoot

            $shareFunctionsPath = "$scriptPath\SharedFunctions.ps1"
            if ($scriptPath.StartsWith("\\")) {
            } else {
              if (!(Test-Path -Path $shareFunctionsPath)) {
                 throw "Missing Dependency File SharedFunctions.ps1"    
              }
            }
            . $shareFunctionsPath

            $mainRegPath = Get-OfficeCTRRegPath
            if (!($mainRegPath)) {
               throw "Office 365 ProPlus is not installed"
            }

            $configRegPath = $mainRegPath + "\Configuration"
            $GPORegPath = "HKLM:\Software\Policies\Microsoft\Office\16.0\common\officeupdate"
            $GPORegPath2 = "Software\Policies\Microsoft\Office\16.0\common\officeupdate"

            $GPOUpdateSource = $true
            $currentUpdateSource = (Get-ItemProperty $GPORegPath -Name updatepath -ErrorAction SilentlyContinue).updatepath

            if(!($currentUpdateSource)){
              $currentUpdateSource = (Get-ItemProperty HKLM:\$configRegPath -Name UpdateUrl -ErrorAction SilentlyContinue).UpdateUrl
              $GPOUpdateSource = $false
            }

            $saveUpdateSource = (Get-ItemProperty HKLM:\$configRegPath -Name SaveUpdateUrl -ErrorAction SilentlyContinue).SaveUpdateUrl
            $clientFolder = (Get-ItemProperty HKLM:\$configRegPath -Name ClientFolder -ErrorAction SilentlyContinue).ClientFolder

            $officeUpdateCDN = Get-OfficeCDNUrl

            $officeCDN = "http://officecdn.microsoft.com"
            $oc2rcFilePath = Join-Path $clientFolder "\OfficeC2RClient.exe"

            $oc2rcParams = "/update user"
            if ($ForceAppShutdown) {
              $oc2rcParams += " forceappshutdown=true"
            } else {
              $oc2rcParams += " forceappshutdown=false"
            }

            if ($UpdatePromptUser) {
              $oc2rcParams += " updatepromptuser=true"
            } else {
              $oc2rcParams += " updatepromptuser=false"
            }

            if ($DisplayLevel) {
              $oc2rcParams += " displaylevel=true"
            } else {
              $oc2rcParams += " displaylevel=false"
            }

            if ($UpdateToVersion) {
              $oc2rcParams += " updatetoversion=$UpdateToVersion"
            }

            [string]$localUpdatePath = ""
            [bool]$scriptPathIsUpdateSource = $false
            if ($UseScriptLocationAsUpdateSource) {
              if ($scriptPath) {
                  if (Test-ItemPathUNC -Path "$scriptPath\SourceFiles") {
                     $localUpdatePath = Change-UpdatePathToChannel -UpdatePath "$scriptPath\SourceFiles" -ValidateUpdateSourceFiles $ValidateUpdateSourceFiles                     
                  } else {
                     $localUpdatePath = Change-UpdatePathToChannel -UpdatePath $scriptPath -ValidateUpdateSourceFiles $ValidateUpdateSourceFiles
                  }

                  [bool]$localIsAlive = Test-UpdateSource -UpdateSource $localUpdatePath -ValidateUpdateSourceFiles $ValidateUpdateSourceFiles

                  if ($localIsAlive) {
                      $scriptPathIsUpdateSource = $true
                      $currentUpdateSource = $localUpdatePath
                  }
              }  
            }

            $UpdateSource = "http"
            if ($currentUpdateSource) {
              If ($currentUpdateSource.StartsWith("\\",1)) {
                 $UpdateSource = "UNC"
              }
            }

            if ($EnableUpdateAnywhere) {

                if ($currentUpdateSource) {
                    [bool]$isAlive = $false
                    if ($currentUpdateSource.ToLower() -eq $officeUpdateCDN.ToLower() -and ($saveUpdateSource)) {
                        if ($currentUpdateSource -ne $saveUpdateSource) {
                            $channelUpdateSource = Change-UpdatePathToChannel -UpdatePath $saveUpdateSource -ValidateUpdateSourceFiles $ValidateUpdateSourceFiles

                            if ($channelUpdateSource -ne $saveUpdateSource) {
                                $saveUpdateSource = $channelUpdateSource
                            }

	                        $isAlive = Test-UpdateSource -UpdateSource $saveUpdateSource -ValidateUpdateSourceFiles $ValidateUpdateSourceFiles
                            if ($isAlive) {
                               Write-Log -Message "Restoring Saved Update Source $saveUpdateSource" -severity 1 -component "Office 365 Update Anywhere"

                               if ($GPOUpdateSource) {
                                   New-ItemProperty -Path "HKLM:\$GPORegPath" -Name "updatepath" -Value $saveUpdateSource -PropertyType String -Force -ErrorAction Stop | Out-Null
                               } else {
                                   New-ItemProperty -Path "HKLM:\$configRegPath" -Name "UpdateUrl" -Value $saveUpdateSource -PropertyType String -Force -ErrorAction Stop | Out-Null
                               }
                            }
                        }
                    }
                }

                if (!($currentUpdateSource)) {
                   if ($officeUpdateCDN) {
                       Write-Log -Message "No Update source is set so defaulting to Office CDN" -severity 1 -component "Office 365 Update Anywhere"

                       if ($GPOUpdateSource) {
                           New-ItemProperty -Path "HKLM:\$GPORegPath" -Name "updatepath" -Value $officeUpdateCDN -PropertyType String -Force -ErrorAction Stop | Out-Null
                       } else {
                           New-ItemProperty -Path "HKLM:\$configRegPath" -Name "UpdateUrl" -Value $officeUpdateCDN -PropertyType String -Force -ErrorAction Stop | Out-Null
                       }

                       $currentUpdateSource = $officeUpdateCDN
                   }
                }

                if (!$isAlive) {
                    $channelUpdateSource = Change-UpdatePathToChannel -UpdatePath $currentUpdateSource -ValidateUpdateSourceFiles $ValidateUpdateSourceFiles

                    if ($channelUpdateSource -ne $currentUpdateSource) {
                        $currentUpdateSource = $channelUpdateSource
                    }

                    $isAlive = Test-UpdateSource -UpdateSource $currentUpdateSource -ValidateUpdateSourceFiles $ValidateUpdateSourceFiles
                    if (!($isAlive)) {
                        if ($currentUpdateSource.ToLower() -ne $officeUpdateCDN.ToLower()) {
                            Set-Reg -Hive "HKLM" -keyPath $configRegPath -ValueName "SaveUpdateUrl" -Value $currentUpdateSource -Type String
                        }

                        Write-Host "Unable to use $currentUpdateSource. Will now use $officeUpdateCDN"
                        Write-Log -Message "Unable to use $currentUpdateSource. Will now use $officeUpdateCDN" -severity 1 -component "Office 365 Update Anywhere"

                        if ($GPOUpdateSource) {
                            New-ItemProperty -Path "HKLM:\$GPORegPath" -Name "updatepath" -Value $officeUpdateCDN -PropertyType String -Force -ErrorAction Stop | Out-Null
                        } else {
                            New-ItemProperty -Path "HKLM:\$configRegPath" -Name "UpdateUrl" -Value $officeUpdateCDN -PropertyType String -Force -ErrorAction Stop | Out-Null
                        }

                        $isAlive = Test-UpdateSource -UpdateSource $officeUpdateCDN -ValidateUpdateSourceFiles $ValidateUpdateSourceFiles
                    }
                }

            } else {
                if($currentUpdateSource -ne $null){
                    $channelUpdateSource = Change-UpdatePathToChannel -UpdatePath $currentUpdateSource -ValidateUpdateSourceFiles $ValidateUpdateSourceFiles

                    if ($channelUpdateSource -ne $currentUpdateSource) {
                        $currentUpdateSource= $channelUpdateSource
                    }

                    $isAlive = Test-UpdateSource -UpdateSource $currentUpdateSource -ValidateUpdateSourceFiles $ValidateUpdateSourceFiles

                }else{
                    $isAlive = Test-UpdateSource -UpdateSource $officeUpdateCDN -ValidateUpdateSourceFiles $ValidateUpdateSourceFiles
                    $currentUpdateSource = $officeUpdateCDN;
                }
            }

            if ($isAlive) {
               if (!($scriptPathIsUpdateSource)) {
                   if ($GPOUpdateSource) {
                     $currentUpdateSource = (Get-ItemProperty $GPORegPath -Name updatepath -ErrorAction SilentlyContinue).updatepath
                   } else {
                     $currentUpdateSource = (Get-ItemProperty HKLM:\$configRegPath -Name UpdateUrl -ErrorAction SilentlyContinue).UpdateUrl
                   }
               }

               if($currentUpdateSource.ToLower().StartsWith("http")){
                   $channelUpdateSource = $currentUpdateSource
               }
               else{
                   $channelUpdateSource = Change-UpdatePathToChannel -UpdatePath $currentUpdateSource -ValidateUpdateSourceFiles $ValidateUpdateSourceFiles
               }

               if (($channelUpdateSource -ne $currentUpdateSource) -or $scriptPathIsUpdateSource) {
                   if ($scriptPathIsUpdateSource) {
                       if ($GPOUpdateSource) {
                          New-ItemProperty -Path "HKLM:\$GPORegPath2" -Name "updatepath" -Value $localUpdatePath -PropertyType String -Force -ErrorAction Stop | Out-Null
                       } else {
                          New-ItemProperty -Path "HKLM:\$configRegPath" -Name "UpdateUrl" -Value $localUpdatePath -PropertyType String -Force -ErrorAction Stop | Out-Null
                       }
                       $channelUpdateSource = $localUpdatePath

                       Remove-ItemProperty HKLM:\$configRegPath -Name SaveUpdateUrl -Force -ErrorAction SilentlyContinue
                       $saveUpdateSource = $null
                   } elseif ($GPOUpdateSource) {
                       New-ItemProperty -Path "HKLM:\$GPORegPath2" -Name "updatepath" -Value $channelUpdateSource -PropertyType String -Force -ErrorAction Stop | Out-Null
                       $channelUpdateSource = $channelUpdateSource
                   } else {
                       New-ItemProperty -Path "HKLM:\$configRegPath" -Name "UpdateUrl" -Value $channelUpdateSource -PropertyType String -Force -ErrorAction Stop | Out-Null
                       $channelUpdateSource = $channelUpdateSource
                   }
               }

               Write-Host "Starting Update process"
               Write-Host "Update Source: $channelUpdateSource" 
               Write-Log -Message "Will now execute $oc2rcFilePath $oc2rcParams with UpdateSource:$channelUpdateSource" -severity 1 -component "Office 365 Update Anywhere"

               StartProcess -execFilePath $oc2rcFilePath -execParams $oc2rcParams

               if ($WaitForUpdateToFinish) {
                    Wait-ForOfficeCTRUpadate
               }

               $saveUpdateSource = (Get-ItemProperty HKLM:\$configRegPath -Name SaveUpdateUrl -ErrorAction SilentlyContinue).SaveUpdateUrl
               if ($saveUpdateSource) {
                   if ($GPOUpdateSource) {
                       New-ItemProperty -Path "HKLM:\$GPORegPath2" -Name "updatepath" -Value $saveUpdateSource -PropertyType String -Force -ErrorAction Stop | Out-Null
                   } else {
                       New-ItemProperty -Path "HKLM:\$configRegPath" -Name "UpdateUrl" -Value $saveUpdateSource -PropertyType String -Force -ErrorAction Stop | Out-Null
                   }
                   Remove-ItemProperty HKLM:\$configRegPath -Name SaveUpdateUrl -Force -ErrorAction SilentlyContinue
               }

            } else {
               $currentUpdateSource = (Get-ItemProperty HKLM:\$configRegPath -Name UpdateUrl -ErrorAction SilentlyContinue).UpdateUrl
               Write-Host "Update Source '$currentUpdateSource' Unavailable"
               Write-Log -Message "Update Source '$currentUpdateSource' Unavailable" -severity 1 -component "Office 365 Update Anywhere"
            }

       } catch {
           Write-Log -Message $_.Exception.Message -severity 1 -component $LogFileName
           throw;
       }
    }
}

Update-Office365Anywhere -WaitForUpdateToFinish $WaitForUpdateToFinish -EnableUpdateAnywhere $EnableUpdateAnywhere -ForceAppShutdown $ForceAppShutdown -UpdatePromptUser $UpdatePromptUser -DisplayLevel $DisplayLevel -UpdateToVersion $UpdateToVersion -LogPath $LogPath -LogName $LogName -ValidateUpdateSourceFiles $ValidateUpdateSourceFiles -UseScriptLocationAsUpdateSource $UseScriptLocationAsUpdateSource



