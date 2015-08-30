[CmdletBinding()]
Param(
    [Parameter()]
    [bool] $WaitForUpdateToFinish = $true,

    [Parameter()]
    [bool] $EnableUpdateAnywhere = $true
)

$os=Get-WMIObject win32_operatingsystem
$osArchitecture = $os.OSArchitecture

if ($osArchitecture -eq "32-bit") {
   $ProgramFiles32Bit = [Environment]::GetFolderPath("ProgramFiles")
} else {
   $ProgramFiles32Bit = [Environment]::GetFolderPath("ProgramFilesx86")
   $ProgramFiles64Bit = [Environment]::GetFolderPath("ProgramFiles")
}

Function Write-Log {
 
    PARAM
	(
         [String]$Message,
         [String]$Path = $LogFilePath,
         [int]$severity,
         [string]$component
	)
 
    try {
        # Get Windows Folder Path
        $windowsDirectory = [Environment]::GetFolderPath("Windows")

        # Build log folder
        $logFolderPath = "$windowsDirectory\CCM\logs"

         # Set log file name
        $LogFileName = "Office365UpdateAnywhere.log"

        # Build log path
        $LogFilePath = Join-Path $logFolderPath $LogFileName

        # Create log file
        If (!($(Test-Path $LogFilePath -PathType Leaf)))
        {
            $null = New-Item -Path $LogFilePath -ItemType File -ErrorAction SilentlyContinue
        }

	    $TimeZoneBias = Get-WmiObject -Query "Select Bias from Win32_TimeZone"
        $Date= Get-Date -Format "HH:mm:ss.fff"
        $Date2= Get-Date -Format "MM-dd-yyyy"
        $type=1
 
        if ($Path) {
           "<![LOG[$Message]LOG]!><time=$([char]34)$date$($TimeZoneBias.bias)$([char]34) date=$([char]34)$date2$([char]34) component=$([char]34)$component$([char]34) context=$([char]34)$([char]34) type=$([char]34)$severity$([char]34) thread=$([char]34)$([char]34) file=$([char]34)$([char]34)>"| Out-File -FilePath $Path -Append -NoClobber -Encoding default
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

        $items = Get-Item . | Select-Object -ExpandProperty property
        $properties = $items | ForEach-Object {
           New-Object psobject -Property @{"property"=$_; "Value" = (Get-ItemProperty -Path . -Name $_).$_}
        }

        $value = $properties | Select Value
        $firstItem = $value[0]
        [string] $cdnPath = $firstItem.Value

        $CDNBaseUrl = Select-String -InputObject $cdnPath -Pattern "http://officecdn.microsoft.com/.*/.{8}-.{4}-.{4}-.{4}-.{12}" -AllMatches | % { $_.Matches } | % { $_.Value }
        Pop-Location
    }
    return $CDNBaseUrl
}

Function Get-OfficeCTRRegPath() {
    $path15 = 'SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration'
    $path16 = 'SOFTWARE\Microsoft\Office\ClickToRun\Configuration'

    if (Test-Path "HKLM:\$path15") {
      return $path15
    } else {
      if (Test-Path "HKLM:\$path16") {
         return $path16
      }
    }
}

Function Get-OfficeCTRScenarioRegPath() {
    $path15 = 'SOFTWARE\Microsoft\Office\15.0\ClickToRun\scenario'
    $path16 = 'SOFTWARE\Microsoft\Office\ClickToRun\scenario'

    if (Test-Path "HKLM:\$path15") {
      return $path15
    } else {
      if (Test-Path "HKLM:\$path16") {
         return $path16
      }
    }
}

Function Test-UpdateSource() {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [string] $UpdateSource = $NULL
    )

  	$uri = [System.Uri]$UpdateSource
	$sourceIsAlive = Test-Connection -Count 1 -computername $uri.Host -Quiet
    return $sourceIsAlive
}

Function Update-Office365Anywhere() {
    [CmdletBinding()]
    Param(
        [Parameter()]
        [bool] $WaitForUpdateToFinish = $true,

        [Parameter()]
        [bool] $EnableUpdateAnywhere = $true
    )

    $officeRegPath = Get-OfficeCTRRegPath

    $currentUpdateSource = (Get-ItemProperty HKLM:\$officeRegPath -Name UpdateUrl -ErrorAction SilentlyContinue).UpdateUrl
    $saveUpdateSource = (Get-ItemProperty HKLM:\$officeRegPath -Name SaveUpdateUrl -ErrorAction SilentlyContinue).SaveUpdateUrl
    $clientFolder = (Get-ItemProperty HKLM:\$officeRegPath -Name ClientFolder -ErrorAction SilentlyContinue).ClientFolder

    $officeUpdateCDN = Get-OfficeCDNUrl

    $officeCDN = "http://officecdn.microsoft.com"
    $oc2rcFilePath = Join-Path $clientFolder "\OfficeC2RClient.exe"
    $oc2rcParams = "/update user forceappshutdown=false updatepromptuser=true displaylevel=true"
    
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
	                $isAlive = Test-UpdateSource -UpdateSource $saveUpdateSource
                    if ($isAlive) {
                       Write-Log -Message "Restoring Saved Update Source $saveUpdateSource" -severity 1 -component "Office 365 Update Anywhere"
                       Set-Reg -Hive "HKLM" -keyPath $officeRegPath -ValueName "UpdateUrl" -Value $saveUpdateSource -Type String
                    }
                }
            }
        }

        if (!($currentUpdateSource)) {
           if ($officeUpdateCDN) {
               Write-Log -Message "No Update source is set so defaulting to Office CDN" -severity 1 -component "Office 365 Update Anywhere"
               Set-Reg -Hive "HKLM" -keyPath $officeRegPath -ValueName "UpdateUrl" -Value $officeUpdateCDN -Type String
               $currentUpdateSource = $officeUpdateCDN
           }
        }

        if (!$isAlive) {
            $isAlive = Test-UpdateSource -UpdateSource $currentUpdateSource
            if (!($isAlive)) {
                       if ($currentUpdateSource.ToLower() -ne $officeUpdateCDN.ToLower()) {
               Set-Reg -Hive "HKLM" -keyPath $officeRegPath -ValueName "SaveUpdateUrl" -Value $currentUpdateSource -Type String
           }

               Write-Log -Message "Unable to use $currentUpdateSource. Will now use $officeUpdateCDN" -severity 1 -component "Office 365 Update Anywhere"
               Set-Reg -Hive "HKLM" -keyPath $officeRegPath -ValueName "UpdateUrl" -Value $officeUpdateCDN -Type String
            }
        }
    } else {
      $isAlive = Test-UpdateSource -UpdateSource $currentUpdateSource
    }

    if ($isAlive) {
       Write-Log -Message "Will now execute $oc2rcFilePath $oc2rcParams" -severity 1 -component "Office 365 Update Anywhere"
       StartProcess -execFilePath $oc2rcFilePath -execParams $oc2rcParams

       if ($WaitForUpdateToFinish) {
            Wait-ForOfficeCTRUpadate
       }
    }

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
       Write-Host "Waiting for Update to Complete..."

       Start-Sleep -Seconds 5

       $scenarioPath = Get-OfficeCTRScenarioRegPath

       $regProv = Get-Wmiobject -list "StdRegProv" -namespace root\default -ErrorAction Stop

       [DateTime]$startTime = Get-Date

       $failure = $false
       $updateRunning=$false
       [string[]]$trackProgress = @()
       [string[]]$trackComplete = @()
       do {
           $allComplete = $true
           
           $scenarioKeys = $regProv.EnumKey($HKLM, $scenarioPath)
           foreach ($scenarioKey in $scenarioKeys.sNames) {
              if ($scenarioKey.ToUpper() -eq "UPDATE") {
                   $taskKeyPath = Join-Path $scenarioPath "$scenarioKey\TasksState"
                   $taskValues = $regProv.EnumValues($HKLM, $taskKeyPath).sNames

                    foreach ($taskValue in $taskValues) {
                        [string]$status = $regProv.GetStringValue($HKLM, $taskKeyPath, $taskValue).sValue
                        $operation = $taskValue.Split(':')[0]
                        $keyValue = $taskValue

                        if ($status.ToUpper() -eq "TASKSTATE_FAILED") {
                          $failure = $true
                        }

                        if (($status.ToUpper() -eq "TASKSTATE_COMPLETED") -or`
                            ($status.ToUpper() -eq "TASKSTATE_CANCELLED") -or`
                            ($status.ToUpper() -eq "TASKSTATE_FAILED")) {
                            if ($trackProgress.Contains($keyValue) -and !$trackComplete.Contains($keyValue)) {
                                $displayValue = $operation + "`t" + $status
                                Write-Host $displayValue
                                $trackComplete += $keyValue 
                            }
                        } else {
                            $allComplete = $false
                            $updateRunning=$true

                            if (!$trackProgress.Contains($keyValue)) {
                                $trackProgress += $keyValue 
                                $displayValue = $operation + "`t" + $status
                                Write-Host $displayValue
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

       if ($updateRunning) {
          if ($failure) {
            Write-Host "Update Failed"
          } else {
            Write-Host "Update Complete"
          }
       } else {
          Write-Host "Update Not Running"
       } 
    }
}

Update-Office365Anywhere -WaitForUpdateToFinish $WaitForUpdateToFinish -EnableUpdateAnywhere $EnableUpdateAnywhere

