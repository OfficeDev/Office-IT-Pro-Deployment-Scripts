Function Check-OfficeUpdateNetworkLoad() {
<#
.SYNOPSIS
Determines the size of the update and quality of delta compression
for an office update.

.DESCRIPTION
Uses Office Deployment Tool to download and install a specified starting
version of Office. Then captures the current received bytes on the 
NetAdapter, before starting an update to the specified end version.
Records the end received bytes to determine the total size of the 
download. Then zips the apply folder within the office updates folder
to determine what the max download size would be without delta
compression. Comparing these two values provides the delta compression
value.

.Notes
Recommended use is running this script on a clean VM.

.PARAMETER VersionStart
The version to install initially before updating

.PARAMETER VersionEnd
The version to update to

.Example
./Check-OfficeUpdateNetworkLoad -VersionStart 15.0.4623.1003 -VersionEnd 15.0.4631.1002
Installs Version 15.04623.1003 and updates to version 15.0.4631.1002 and returns the 
network traffic numbers. (In original test environment this call returned the values
MaxDownload: ~324000000, ActualDownload: ~128500000, DeltaCompressionRate: ~0.60)

.Outputs
Hashtable with values for Downloaded bytes, max size, delta compression rate

#>

Param(
    [Parameter(Mandatory=$true)]
    [string] $VersionStart,

    [Parameter(Mandatory=$true)]
    [string] $VersionEnd,

    [Parameter()]
    [string] $SourcePath
)

Begin{
$ZipPath = "$env:USERPROFILE\Downloads\sizeTest.zip"

if ($SourcePath) {
    $config1 = 
    "<Configuration>
        <Add OfficeClientEdition=`"32`" Version=`"$VersionStart`" SourcePath=`"$SourcePath`" >
            <Product ID=`"O365ProPlusRetail`">
                <Language ID=`"en-us`" />
            </Product>
        </Add>
        <Updates Enabled=`"FALSE`" UpdatePath=`"$SourcePath`" />
    </Configuration>"
} else {
    $config1 = 
    "<Configuration>
        <Add OfficeClientEdition=`"32`" Version=`"$VersionStart`" >
            <Product ID=`"O365ProPlusRetail`">
                <Language ID=`"en-us`" />
            </Product>
        </Add>
        <Updates Enabled=`"FALSE`" UpdatePath=`"$SourcePath`" />
    </Configuration>"
}

$folderPath = "$env:ProgramFiles\Microsoft Office 15\Data\Updates\Apply"
$ODTSource = "http://download.microsoft.com/download/6/2/3/6230F7A2-D8A9-478B-AC5C-57091B632FCF/officedeploymenttool_x86_4747-1000.exe"
}

Process{


if (!(Test-Path -Path "$PSScriptRoot\Setup.exe")) {
    #download setup
    Invoke-WebRequest $ODTSource -OutFile "$env:USERPROFILE\Downloads\officedeploymenttool_x86_4747-1000.exe" | Out-Null
    Set-Location "$env:USERPROFILE\Downloads"
    .\officedeploymenttool_x86_4747-1000.exe /extract:$env:USERPROFILE\downloads\ODT /passive /quiet | Out-Null
    Set-Location ODT
}


#build configuration file
$config1 | Out-File  "$PSScriptRoot\configuration.xml"
Invoke-Expression -Command "$PSScriptRoot\Setup.exe /configure configuration.xml" | Out-Null

$mainRegPath = Get-OfficeCTRRegPath
$configRegPath = $mainRegPath + "\Configuration"
$clientFolder = (Get-ItemProperty HKLM:\$configRegPath -Name ClientFolder -ErrorAction SilentlyContinue).ClientFolder
$installationPath = (Get-ItemProperty HKLM:\$configRegPath -Name InstallationPath -ErrorAction SilentlyContinue).InstallationPath

#Start word to block update from applying when finished downloading
$winWordPath = ""
$winWordPath15 = $installationPath + "\root\office15\WINWORD.EXE"
$winWordPath16 = $installationPath + "\root\Office16\WINWORD.EXE"
if (Test-Path -Path $winWordPath15) {
  $winWordPath = $winWordPath15
}
if (Test-Path -Path $winWordPath16) {
  $winWordPath = $winWordPath16
}

Start-Process $winWordPath

Wait-ForOfficeCTRUpadate

#get bytes for net adapter
$netstat1 = Get-NetAdapterStatistics

#Start update
Start-Process "$clientFolder\OfficeC2RClient.exe" "/update user updatetoversion=$VersionEnd"

#Wait for update to complete and stop the UAC process if it gets in the way
$complete = $false
while($complete -eq $false){
    $procs = Get-Process | ? ProcessName -eq 'officeclicktorun'
    $UACProc = Get-Process | ? ProcessName -eq "consent"
    if($UACProc -ne $null){
        $UACProc.Kill()
        $UACProc = $null
        $complete = $true
    }
    foreach($proc in $procs){
        if($proc.MainWindowTitle -eq "Please close programs" -or $proc.MainWindowTitle -eq "We need to close some programs"){
            $complete = $true
        }
    }
}

#get bytes for net adapter
$netstat2 = Get-NetAdapterStatistics
$bytes = 0
if($netstat1.GetType() -is [array]){
    foreach($adapter in $netstat2){
        $bytes += $adapter.ReceivedBytes
    }
    foreach($adapter in $netstat1){
        $bytes -= $adapter.ReceivedBytes
    }
}else{
    $bytes = $netstat2.ReceivedBytes - $netstat1.ReceivedBytes 
}

#Zip the Data/Updates/Apply folder to get what size of update could have been
Add-Type -assembly "system.io.compression.filesystem"

if (!(Test-Path -Path $folderPath)) {
  $folderPath = "$installationPath\Updates\Apply"
}

if (!(Test-Path -Path $folderPath)) {
  $folderPath = "$installationPath\Data\Updates\Apply"
}

[io.compression.zipfile]::CreateFromDirectory($folderPath, $ZipPath)
$zipSize = Get-Item $ZipPath

#Stop word process
$word = Get-Process | ? ProcessName -eq WINWORD
$word.Kill()
$word = $null

#Output results
@{
    ActualDownload = $bytes;
    MaxDownload = $zipSize.Length;
    DeltaCompressionRate = 1 - ($bytes/$zipSize.Length);
}
}
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
                        if ($trackProgress.Contains($keyValue) -and !$trackComplete.Contains($keyValue)) {
                            $displayValue = $operation + "`t" + $status + "`t" + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
                            #Write-Host $displayValue
                            $trackComplete += $keyValue 

                            $statusName = $status.Split('_')[1];

                            if ($operation.ToUpper().Contains("DOWNLOAD") -or `
                                $operation.ToUpper().Contains("APPLY")) {

                                $operationTime = getOperationTime -OperationStart $operationStart

                                $displayText = $statusName + "`t" + $operationTime

                                Write-Host $displayText
                            }
                        }
                    } else {
                        $allComplete = $false
                        $updateRunning=$true


                        if (!$trackProgress.Contains($keyValue)) {
                             $trackProgress += $keyValue 
                             $displayValue = $operation + "`t" + $status + "`t" + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')

                             $operationStart = Get-Date

                             if ($operation.ToUpper().Contains("DOWNLOAD")) {
                                Write-Host "Downloading Update: " -NoNewline
                             }

                             if ($operation.ToUpper().Contains("APPLY")) {
                                Write-Host "Applying Update: " -NoNewline
                             }

                             if ($operation.ToUpper().Contains("FINALIZE")) {
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
