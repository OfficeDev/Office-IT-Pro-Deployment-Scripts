param(
    [Parameter()]
    [ValidateSet("FirstReleaseCurrent","Current","FirstReleaseDeferred","Deferred",
                 "MonthlyTargeted","Monthly","SemiAnnualTargeted","SemiAnnual")]
    [string]$Channel,

    [Parameter()]
    [switch]$RollBack,

    [Parameter()]
    [bool]$SendExitCode = $false,

    [Parameter()]
    [string]$LogFilePath
)

Function Get-ScriptPath() {
  [CmdletBinding()]
  param(

  )

  process {
    #get local path
    $scriptPath = "."

    if ($PSScriptRoot) {
        $scriptPath = $PSScriptRoot
    } else {
        $scriptPath = (Get-Item -Path ".\").FullName
    }
    return $scriptPath
  }
}

Function Get-OfficeC2Rexe() {
    [CmdletBinding()]
    Param(

    )
    process {
        $Office2RClientKey = 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' #ClientFolder

        #find update exe file
        $OfficeUpdatePath = Get-ItemProperty -Path $Office2RClientKey | Select-Object -Property ClientFolder
        $temp = Out-String -InputObject $OfficeUpdatePath
        $temp = $temp.Substring($temp.LastIndexOf('-')+2)
        $temp = $temp.Trim()
        $OfficeUpdatePath = $temp
        $OfficeUpdatePath+= '\OfficeC2RClient.exe'
        return $OfficeUpdatePath
    }
}

Function Wait-ForOfficeCTRUpdate() {
    [CmdletBinding()]
    Param(
        [Parameter()]
        [int] $TimeOutInMinutes = 120
    )

    begin {
        $HKLM = [UInt32] "0x80000002"
        $HKCR = [UInt32] "0x80000000"

        $currentFileName = Get-CurrentFileName
        Set-Alias -name LINENUM -value Get-CurrentLineNumber
    }

    process {
       Write-Host "Waiting for Update process to Complete..."
       WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Waiting for Update process to Complete..." -LogFilePath $LogFilePath

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
                                WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError $displayText -LogFilePath $LogFilePath
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
                                WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Downloading Update: " -LogFilePath $LogFilePath
                             }

                             if ($operation.ToUpper().IndexOf("APPLY") -gt -1) {
                                Write-Host "Applying Update: " -NoNewline
                                WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Applying Update: " -LogFilePath $LogFilePath
                             }

                             if ($operation.ToUpper().IndexOf("FINALIZE") -gt -1) {
                                Write-Host "Finalizing Update: " -NoNewline
                                WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Applying Update: " -LogFilePath $LogFilePath
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
       WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError $displayValue -LogFilePath $LogFilePath

       $totalOperationTime = getOperationTime -OperationStart $totalOperationStart
       [bool]$UpdateCompleted = $true

       if ($updateRunning) {
          if ($failure) {
            $UpdateCompleted = $false
            Write-Host "Update Failed"
            WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Update Failed" -LogFilePath $LogFilePath
            throw "Update Failed"
          } else {
            Write-Host "Update Completed - Total Time: $totalOperationTime"
            WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Update Completed - Total Time: $totalOperationTime" -LogFilePath $LogFilePath
          }
       } else {
            $UpdateCompleted = $false
            Write-Host "Update Not Running"
            WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Update Not Running" -LogFilePath $LogFilePath
       }

       return $UpdateCompleted
    
    }
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

Function Test-UpdateSource() {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [string] $UpdateSource = $NULL,

        [Parameter()]
        [string]$LogFilePath
    )
    
    $currentFileName = Get-CurrentFileName
    Set-Alias -name LINENUM -value Get-CurrentLineNumber

  	$uri = [System.Uri]$UpdateSource

    [bool]$sourceIsAlive = $false

    if($uri.Host){
	    $sourceIsAlive = Test-Connection -Count 1 -computername $uri.Host -Quiet
    }else{
        $sourceIsAlive = Test-Path $uri.LocalPath -ErrorAction SilentlyContinue
    }

    if ($sourceIsAlive) {
        $sourceIsAlive = Validate-UpdateSource -UpdateSource $UpdateSource
    }

    WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "sourceIsAlive set to $sourceIsAlive" -LogFilePath $LogFilePath

    return $sourceIsAlive
}

Function Test-Url() {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [string] $Url = $NULL
    )

# First we create the request.
$HTTP_Request = [System.Net.WebRequest]::Create($Url)

# We then get a response from the site.
$HTTP_Response = $HTTP_Request.GetResponse()

# We then get the HTTP code as an integer.
$HTTP_Status = [int]$HTTP_Response.StatusCode

# Finally, we clean up the http request by closing it.
$HTTP_Response.Close()

If ($HTTP_Status -eq 200) { 
    return $true
}
Else {
    return $false
}
}

function Test-ItemPathUNC() {    [CmdletBinding()]	
    Param
	(	    [Parameter(Mandatory=$true)]
	    [String]$Path,	    [Parameter()]
	    [String]$FileName = $null    )    Process {       $pathExists = $false       if ($FileName) {         $filePath = "$Path\$FileName"         $pathExists = [System.IO.File]::Exists($filePath)       } else {         $pathExists = [System.IO.Directory]::Exists($Path)         if (!($pathExists)) {            $pathExists = [System.IO.File]::Exists($Path)         }       }       return $pathExists;    }}

Function Validate-UpdateSource() {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [string] $UpdateSource = $NULL,

        [Parameter()]
        [string] $OfficeClientEdition,
        
        [Parameter()]
        [string] $Bitness = "x86",

        [Parameter()]
        [string[]] $OfficeLanguages = $null,

        [Parameter()]
        [bool]$ShowMissingFiles = $true,

        [Parameter()]
        [string]$LogFilePath
    )

    Set-Alias -name LINENUM -value Get-CurrentLineNumber
    $currentFileName = Get-CurrentFileName

    if(!$OfficeClientEdition)
    {
        #checking if office client edition is null, if not, set bitness to client office edition
    }
    else
    {
        $Bitness = $OfficeClientEdition
    }

    [bool]$validUpdateSource = $true
    [string]$cabPath = ""

    if ($UpdateSource) {
        $mainRegPath = Get-OfficeCTRRegPath
        if ($mainRegPath) {
            $configRegPath = $mainRegPath + "\Configuration"
            $currentplatform = (Get-ItemProperty HKLM:\$configRegPath -Name Platform -ErrorAction SilentlyContinue).Platform
            $updateToVersion = (Get-ItemProperty HKLM:\$configRegPath -Name UpdateToVersion -ErrorAction SilentlyContinue).UpdateToVersion
            $llcc = (Get-ItemProperty HKLM:\$configRegPath -Name ClientCulture -ErrorAction SilentlyContinue).ClientCulture
        }

        $currentplatform = $Bitness

        $mainCab = "$UpdateSource\Office\Data\v32.cab"
        $bitness = "32"
        if ($currentplatform -eq "x64") {
            $mainCab = "$UpdateSource\Office\Data\v64.cab"
            $bitness = "64"
        }

        if (!($updateToVersion)) {
           $cabXml = Get-CabVersion -FilePath $mainCab
           if ($cabXml) {
               $updateToVersion = $cabXml.Version.Available.Build
           }
        }

        [xml]$xml = Get-ChannelXml
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

           if (!($updateFileExists) -and ($checkFile.relativePath -notmatch "Experiment")) {
              $fileExists = $missingFiles.Contains($fullPath)
              if (!($fileExists)) {
                 $missingFiles.Add($fullPath)
                 if($ShowMissingFiles){
                    Write-Host "Source File Missing: $fullPath"
                    WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Source File Missing: $fullPath" -LogFilePath $LogFilePath
                 }
                 Write-Log -Message "Source File Missing: $fullPath" -severity 1 -component "Office 365 Update Anywhere" 
              }     
              $validUpdateSource = $false
           }
        }

    }
    WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "validUpdateSource set to $validUpdateSource" -LogFilePath $LogFilePath

    return $validUpdateSource
}

Function Get-LatestVersion() {
  [CmdletBinding()]
  Param(
     [Parameter(Mandatory=$true)]
     [string] $UpdateURLPath,

     [Parameter()]
     [string]$LogFilePath
  )

  process {
    $currentFileName = Get-CurrentFileName
    Set-Alias -name LINENUM -value Get-CurrentLineNumber

    [array]$totalVersion = @()
    $Version = $null

    $isUrl = $UpdateURLPath -like 'http*'

    $tempUpdateURLPath = "$UpdateURLPath/Office/Data/v32.cab"

    if ($isUrl) {
        $cabXml = Get-UrlCabXml -UpdateURLPath $tempUpdateURLPath
        if ($cabXml) {
            $availNode = $cabXml.Version.Available
            $currentVersion = $availNode.Build
            if ($currentVersion) {
               $Version = $currentVersion
            }
        }
    } else {
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
    }

    WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Latest Version set to $Version" -LogFilePath $LogFilePath

    return $Version
  }
}

Function Get-PreviousVersion() {
  [CmdletBinding()]
  Param(
     [Parameter(Mandatory=$true)]
     [string] $UpdateURLPath,

     [Parameter()]
     [string]$LogFilePath
  )

  process {
    $currentFileName = Get-CurrentFileName
    Set-Alias -name LINENUM -value Get-CurrentLineNumber

    [array]$totalVersion = @()
    $Version = $null

    $LatestBranchVersionPath = $UpdateURLPath + '\Office\Data'
    if(Test-Path $LatestBranchVersionPath){
        $DirectoryList = Get-ChildItem $LatestBranchVersionPath
        Foreach($listItem in $DirectoryList){
            if($listItem.GetType().Name -eq 'DirectoryInfo'){
              if ($listItem.Name -match '\d{2}\.\d\.\d{4}\.\d{4}') {
                $totalVersion+=$listItem.Name
              }
            }
        }
    }

    $totalVersion = $totalVersion | Sort-Object -Descending
    
    #sets version number to the newest version in directory for channel if version is not set by user in argument  
    if($totalVersion.Count -gt 1){
        $Version = $totalVersion[1]
    } else {
        return $null
    } 

    WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Previous Version set to $Version" -LogFilePath $LogFilePath

    return $Version
  }
}

function Change-UpdatePathToChannel {
   [CmdletBinding()]
   param( 
     [Parameter()]
     [string] $UpdatePath,
     
     [Parameter()]
     [string] $Channel,

     [Parameter()]
     [string]$LogFilePath
   )

   $currentFileName = Get-CurrentFileName
   Set-Alias -name LINENUM -value Get-CurrentLineNumber

   $newUpdatePath = $UpdatePath

   $branchShortName = "DC"
   if ($Channel.ToString().ToLower() -eq "current") {
      $branchShortName = "CC"
   }
   if ($Channel.ToString().ToLower() -eq "firstreleasecurrent") {
      $branchShortName = "FRCC"
   }
   if ($Channel.ToString().ToLower() -eq "firstreleasedeferred") {
      $branchShortName = "FRDC"
   }
   if ($Channel.ToString().ToLower() -eq "deferred") {
      $branchShortName = "DC"
   }
   if ($Channel.ToString().ToLower() -eq "monthlytargeted") {
      $branchShortName = "MTC"
   }
   if ($Channel.ToString().ToLower() -eq "monthly") {
      $branchShortName = "MC"
   }
   if ($Channel.ToString().ToLower() -eq "semiannualtargeted") {
      $branchShortName = "SATC"
   }
   if ($Channel.ToString().ToLower() -eq "semiannual") {
      $branchShortName = "SAC"
   }

   $channelNames = @("FRCC", "CC", "FRDC", "DC", "MTC", "MC", "SATC", "SAC")

   $madeChange = $false
   foreach ($channelName in $channelNames) {
      if ($UpdatePath.ToUpper().EndsWith("\$channelName")) {
         $newUpdatePath = $newUpdatePath -replace "\\$channelName", "\$branchShortName"
         $madeChange = $true
      } 
      if ($UpdatePath.ToUpper().Contains("\$channelName\")) {
         $newUpdatePath = $newUpdatePath -replace "\\$channelName\\", "\$branchShortName\"
         $madeChange = $true
      } 
      if ($UpdatePath.ToUpper().EndsWith("/$channelName")) {
         $newUpdatePath = $newUpdatePath -replace "\/$channelName", "/$branchShortName"
         $madeChange = $true
      }
      if ($UpdatePath.ToUpper().Contains("/$channelName/")) {
         $newUpdatePath = $newUpdatePath -replace "\/$channelName\/", "/$branchShortName/"
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

   try {
     $pathAlive = Test-UpdateSource -UpdateSource $newUpdatePath
   } catch {
     $pathAlive = $false
   }
   
   if ($pathAlive) {
     WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "newUpdatePath set to $newUpdatePath" -LogFilePath $LogFilePath
     return $newUpdatePath
   } else {
     WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "UpdatePath set to $UpdatePath" -LogFilePath $LogFilePath
     return $UpdatePath
   }
}

function Test-UpdateSourceTcpPort {
    Param(
        [parameter(ParameterSetName='URL', Position=0)]
        [string]
        $URL,

        [parameter(ParameterSetName='IP', Position=0)]
        [System.Net.IPAddress]
        $IPAddress,

        [parameter(Mandatory=$true , Position=1)]
        [int]
        $Port,

        [parameter()]
        [string]$UpdateSource = $null,

        [Parameter()]
        [string]$LogFilePath
    )

    $currentFileName = Get-CurrentFileName
    Set-Alias -name LINENUM -value Get-CurrentLineNumber

    $sourceIsAlive = $false

    $RemoteServer = If ([string]::IsNullOrEmpty($URL)) {$IPAddress} Else {$URL};

    $test = New-Object System.Net.Sockets.TcpClient;

    Try
    {
        $test.Connect($RemoteServer, $Port);
        $sourceIsAlive = $true
    } Catch {}

    Finally
    {
        $test.Close();
    }

    if ($sourceIsAlive) {
        $sourceIsAlive = Validate-UpdateSource -UpdateSource $UpdateSource
    }

    WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "sourceIsAlive set to $sourceIsAlive" -LogFilePath $LogFilePath

    return $sourceIsAlive
}

function Detect-Channel {
   param( 
        [Parameter()]
        [string]$LogFilePath
   )

Process {
   $currentFileName = Get-CurrentFileName
   Set-Alias -name LINENUM -value Get-CurrentLineNumber 
        
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

   if($CurrentChannel){
      if($CurrentChannel.GetType().Name -eq "Object[]"){
         $CurrentChannel = $CurrentChannel | ? {$_.branch -ne "FirstReleaseCurrent" -and $_.branch -ne "Current" `
                                                                                    -and $_.branch -ne "FirstReleaseDeferred" `
                                                                                    -and $_.branch -ne "Deferred"}
      }
   }

   return $CurrentChannel
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

function Get-ChannelUrl() {
   [CmdletBinding()]
   param( 
      [Parameter(Mandatory=$true)]
      [string]$Channel
   )

   Process {
      $channelXml = Get-ChannelXml

      $currentChannel = $channelXml.UpdateFiles.baseURL | Where {$_.branch -eq $Channel.ToString() }
      return $currentChannel
   }
}

function Get-UrlCabXml() {
   [CmdletBinding()]
   Param(
     [Parameter(Mandatory=$true)]
     [string] $UpdateURLPath
   )

   process {
       $webclient = New-Object System.Net.WebClient
       $XMLFilePath = "$env:TEMP/v32.cab"
       $XMLDownloadURL = $UpdateURLPath
       $webclient.DownloadFile($XMLDownloadURL,$XMLFilePath)

       $tmpName = "VersionDescriptor.xml"
       expand $XMLFilePath $env:TEMP -f:$tmpName | Out-Null
       $tmpName = $env:TEMP + "\VersionDescriptor.xml"
       [xml]$channelXml = Get-Content $tmpName

       return $channelXml
   }

}

function Get-ChannelXml() {
   [CmdletBinding()]
   param( 
      
   )

   process {
       $XMLFilePath = "$PSScriptRoot\ofl.cab"

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
           [System.XML.XMLDocument]$channelXml = Get-Content $tmpName
       }else {
           $scriptPath = Get-ScriptPath
           $tmpName = $scriptPath + "\o365client_64bit.xml"
           [System.XML.XMLDocument]$channelXml = Get-Content $tmpName
       }

       return $channelXml
   }

}

Function Set-OfficeCDNUrl() {
   [CmdletBinding()]
   param( 
      [Parameter(Mandatory=$true)]
      [string]$Channel,

      [Parameter()]
      [string]$LogFilePath
   )

   Process {
        $currentFileName = Get-CurrentFileName
        Set-Alias -name LINENUM -value Get-CurrentLineNumber 

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
        WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "CDNBaseUrl registry key set to $ChannelUrl.URL" -LogFilePath $LogFilePath
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

Function Get-InstalledLanguages() {
    [CmdletBinding()]
    Param(
        [string]$computer = $env:COMPUTERNAME
    )
    process {
       $returnLangs = @()
       $mainRegPath = Get-OfficeCTRRegPath

       if ($mainRegPath) {
          if (Test-Path -Path "hklm:\$mainRegPath\ProductReleaseIDs") {
               $activeConfig = Get-ItemProperty -Path "hklm:\$mainRegPath\ProductReleaseIDs"
               if($activeConfig.ActiveConfiguration){
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
               } else {
                  $HKLM = [UInt32] "0x80000002"
                  $regProv = Get-Wmiobject -list "StdRegProv" -Namespace root\default -ComputerName $computer

                  $activeConfig = "hklm:\$mainRegPath\ProductReleaseIDs"
                  $activeItems = Get-ChildItem -Path $activeConfig
    
                  foreach($config in $activeItems){
                      $item = $config.Name | Split-Path -Leaf
                      $path = Join-Path $activeConfig $item
                  
                      $pathItems = Get-ChildItem -Path $path
                  
                      foreach($pathItem in $pathItems){
                          if($pathItem.Name -match "Culture"){
                              $activeID = $item
                          }
                      }
                  }

                  $languages = (Get-Item -Path "hklm:\$mainRegPath\ProductReleaseIDs\$activeId\culture").Property

                  foreach ($language in $languages) {
                      if ($language.Contains(".")) {
                          $language = $keyName.Split(".")[0]
                      }
                      
                      if ($language.ToLower() -ne "x-none") {
                         $culture = New-Object system.globalization.cultureinfo($language)
                         $returnLangs += $culture
                      }
                  }
               }
          }
       }

       return $returnLangs
    }
}

function Get-CurrentLineNumber {
    $MyInvocation.ScriptLineNumber
}

function Get-CurrentFileName{
    $MyInvocation.ScriptName.Substring($MyInvocation.ScriptName.LastIndexOf("\")+1)
}

Function WriteToLogFile() {
    param( 
        [Parameter(Mandatory=$true)]
        [string]$LNumber,

        [Parameter(Mandatory=$true)]
        [string]$FName,

        [Parameter(Mandatory=$true)]
        [string]$ActionError,

        [Parameter()]
        [string]$LogFilePath
    )

    try{
        $headerString = "Time".PadRight(30, ' ') + "Line Number".PadRight(15,' ') + "FileName".PadRight(60,' ') + "Action"
        $stringToWrite = $(Get-Date -Format G).PadRight(30, ' ') + $($LNumber).PadRight(15, ' ') + $($FName).PadRight(60,' ') + $ActionError

        if(!$LogFilePath){
            $LogFilePath = "$env:windir\Temp\" + (Get-Date -Format u).Substring(0,10)+"_OfficeDeploymentLog.txt"
        }
        if(Test-Path $LogFilePath){
             Add-Content $LogFilePath $stringToWrite
        }
        else{#if not exists, create new
             Add-Content $LogFilePath $headerString
             Add-Content $LogFilePath $stringToWrite
        }
    } catch [Exception]{
        Write-Host $_
    }
}

try {
    $currentFileName = Get-CurrentFileName
    Set-Alias -name LINENUM -value Get-CurrentLineNumber

    if (!($RollBack)) {
      if (!($Channel)) {
         WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Channel Parameter is required. Use the -Channel parameter and enter either Insiders, Monthly, Targeted, or Broad." -LogFilePath $LogFilePath
         throw "Channel Parameter is required. Use the -Channel parameter and enter either Insiders, Monthly, Targeted, or Broad."
      }
    }

    [bool]$PolicyPath = $true
    [bool]$SetBack = $false

    $Office2RClientKey = 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' #ClientFolder
    $OfficePolicyPath = 'HKLM:\Software\Policies\Microsoft\Office\16.0\common\officeupdate'

    $UpdateUrl = (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration -Name UpdateUrl -ErrorAction SilentlyContinue).UpdateUrl 
    $GPOUpdatePath = (Get-ItemProperty HKLM:\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate -Name updatepath -ErrorAction SilentlyContinue).updatepath
    $UpdateChannel = (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration -Name UpdateChannel -ErrorAction SilentlyContinue).UpdateChannel
    
    if($UpdateUrl -ne $NULL){
        $UpdateURLPath = $UpdateUrl
        $PolicyPath = $false
    }
    
    if($GPOUpdatePath -ne $NULL){
        $UpdateURLPath = $GPOUpdatePath
    } 

    if($UpdateUrl -eq $null -and $UpdateChannel -ne $NULL){
        if($GPOUpdatePath -eq $NULL){
            $UpdateURLPath = $UpdateChannel
            $PolicyPath = $false
        } else {
            $UpdateURLPath = $UpdateChannel    
        }
    }
   
    if (!($UpdateURLPath)) {
        $UpdateURLPath = Get-ScriptPath
        $SetBack = $true

        $TmpUpdateUrlPath = "$UpdateURLPath\SourceFiles"
        if (Test-Path -Path $TmpUpdateUrlPath) {
           $UpdateURLPath = $TmpUpdateUrlPath
        } else {
            $UpdateURLPath = (Detect-Channel).URL
            $PolicyPath = $false
            $SetBack = $false
        }
    } else {
        if($UpdateURLPath -notlike '*officecdn.microsoft.com*'){
            $urlPathChk = Test-Path $UpdateURLPath
            if(!($urlPathChk)){
                $UpdateURLPath = Get-ScriptPath
                $SetBack = $true

                $TmpUpdateUrlPath = "$UpdateURLPath\SourceFiles"
                if (Test-Path -Path $TmpUpdateUrlPath) {
                   $UpdateURLPath = $TmpUpdateUrlPath
                }
            }
        }
    }

    WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "UpdateURLPath set to $UpdateURLPath" -LogFilePath $LogFilePath
    $OldUpdatePath = $UpdateURLPath

    $detectChannelUrl = $NULL
    $detectChannel = (Detect-Channel)

    if ($detectChannel) {
        $detectChannelBranch = $detectChannel.Branch
        $detectChannelUrl = $detectChannel.Url
        $oldChannel = $detectChannelBranch
    }

    if ($RollBack) {
       $Channel = $detectChannelBranch
    }

    [bool]$updateUrlIsCdn = $false
    if ($OldUpdatePath) {
      if ($OldUpdatePath -like '*officecdn.microsoft.com*') {
          $updateUrlIsCdn = $true
      }
    }

    if ($updateUrlIsCdn) {
      $newChannelUrl = Get-ChannelUrl -Channel $Channel
      if ($newChannelUrl) {
         $UpdateURLPath = $newChannelUrl.Url
      }
    } else {
      $UpdateURLPath = Change-UpdatePathToChannel -Channel $Channel -UpdatePath $UpdateURLPath
    }
    WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "UpdateURLPath set to $UpdateURLPath" -LogFilePath $LogFilePath
    if($UpdateURLPath -like '*officecdn.microsoft.com*'){
        $validSource = Test-UpdateSourceTcpPort -URL "officecdn.microsoft.com" -Port 80 -UpdateSource $UpdateURLPath
    } else {
        $validSource = Test-UpdateSource -UpdateSource $UpdateURLPath
    }

    if (!($validSource)) {
        WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "UpdateSource not Valid $UpdateURLPath" -LogFilePath $LogFilePath
        throw "UpdateSource not Valid $UpdateURLPath"
    }

    $currentVersion = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration").VersionToReport

    if ($SetBack) {
        $oldUpdatePath = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration").UpdateUrl
        if ($oldUpdatePath) {
            New-ItemProperty $Office2RClientKey -Name BackupUpdateUrl -PropertyType String -Value $oldUpdatePath -Force | Out-Null
            WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "BackupUpdateUrl registry key set to $oldUpdatePath" -LogFilePath $LogFilePath
        }
    }

    if ($UpdateURLPath -and $UpdateUrl -ne $NULL) {
        if ($PolicyPath) {
            New-ItemProperty $OfficePolicyPath -Name updatepath -PropertyType String -Value $UpdateURLPath -Force | Out-Null
            WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "updatepath registry key set to $UpdateURLPath" -LogFilePath $LogFilePath
        } elseif($oldUpdatePath) {
            New-ItemProperty $Office2RClientKey -Name UpdateUrl -PropertyType String -Value $UpdateURLPath -Force | Out-Null
            WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "UpdateUrl registry key set to $UpdateURLPath" -LogFilePath $LogFilePath
        }
    }

    $OfficeUpdatePath = Get-OfficeC2Rexe
    if (!($OfficeUpdatePath)) {
        WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Cannot find OfficeC2RClient.exe file" -LogFilePath $LogFilePath
        throw "Cannot find OfficeC2RClient.exe file"
    }
    
    if ($RollBack) {
      $Version = Get-PreviousVersion -UpdateURLPath $UpdateURLPath
      if (!($Version)) {
        WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Rollback Version Not Available" -LogFilePath $LogFilePath
        throw "Rollback Version Not Available"
      }
    } else {
      $Version = Get-LatestVersion -UpdateURLPath $UpdateURLPath
    }

    if (!($RollBack)) {
           Set-OfficeCDNUrl -Channel $Channel

           if($UpdateChannel -ne $NULL){
               New-ItemProperty $Office2RClientKey -Name UpdateChannel -PropertyType String -Value $UpdateURLPath -Force | Out-Null
               WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "UpdateChannel registry key set to $UpdateURLPath" -LogFilePath $LogFilePath
           }
        }
    
    [bool]$continue = $true
    if(!$RollBack){
        if (($Version) -and ($oldChannel -eq $Channel)) {
            [bool]$continue = $false
            
            Write-Host "The client already has version installed: $Version"
            WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "The client already has version installed: $Version" -LogFilePath $LogFilePath

            Set-OfficeCDNUrl -Channel $Channel
        }
    } else {
        if($Version -ne $currentVersion){ 
            [bool]$continue = $false
        }
    }

    if($continue){
        $arguments = "/update user displaylevel=false forceappshutdown=true updatepromptuser=false updatetoversion=$Version"
        WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "arguments set to /update user displaylevel=false forceappshutdown=true updatepromptuser=false updatetoversion=$Version" -LogFilePath $LogFilePath
                
        #run update exe file
        WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Starting the update" -LogFilePath $LogFilePath
        Start-Process -FilePath $OfficeUpdatePath -ArgumentList $arguments
     
        $UpdateStatus = Wait-ForOfficeCTRUpdate

        if ($UpdateStatus -eq $false){
            Set-OfficeCDNUrl -Channel $oldChannel

            if ($PolicyPath) {
                New-ItemProperty $OfficePolicyPath -Name updatepath -PropertyType String -Value $OldUpdatePath -Force | Out-Null
                WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "updatepath registry key set to $OldUpdatePath" -LogFilePath $LogFilePath
            } elseif($oldUpdatePath) {
                if($UpdateUrl -ne $NULL){
                    New-ItemProperty $Office2RClientKey -Name UpdateUrl -PropertyType String -Value $OldUpdatePath -Force | Out-Null
                    WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "UpdateUrl registry key set to $OldUpdatePath" -LogFilePath $LogFilePath
                }
            }
        }

        if ($SetBack) {
            if ($oldUpdatePath) {
                Remove-ItemProperty $Office2RClientKey -Name BackupUpdateUrl -Force | Out-Null
                WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Removed the BackupUpdateUrl registry key" -LogFilePath $LogFilePath
            }
        }
    } else {
        Write-Host "The channel has been changed to $Channel"
        WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "The channel has been changed to $Channel" -LogFilePath $LogFilePath

        Remove-ItemProperty $Office2RClientKey -Name BackupUpdateUrl -Force -ErrorAction SilentlyContinue | Out-Null
        WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Removed the BackupUpdateUrl registry key" -LogFilePath $LogFilePath
    }

    if ($SendExitCode) {
       [System.Environment]::Exit(0)
    }
} catch {
    Write-Host $_ -ForegroundColor Red
    WriteToLogFile -LNumber $_.InvocationInfo.ScriptLineNumber -FName $currentFileName -ActionError $_
    $Error = $null
    if ($SendExitCode) {
        [System.Environment]::Exit(1)
    }
}

