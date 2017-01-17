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
    [string] $LogPath = $null,

    [Parameter()]
    [string] $LogName = $null,

    [Parameter()]
    [bool] $ValidateUpdateSourceFiles = $true

)

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
        $fileName = $_.InvocationInfo.ScriptName.Substring($_.InvocationInfo.ScriptName.LastIndexOf("\")+1)
        WriteToLogFile -LNumber $_.InvocationInfo.ScriptLineNumber -FName $fileName -ActionError $_
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
        $fileName = $_.InvocationInfo.ScriptName.Substring($_.InvocationInfo.ScriptName.LastIndexOf("\")+1)
        WriteToLogFile -LNumber $_.InvocationInfo.ScriptLineNumber -FName $fileName -ActionError $_
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

Function Test-UpdateSource() {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [string] $UpdateSource = $NULL,

        [Parameter()]
        [bool] $ValidateUpdateSourceFiles = $true
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
           $sourceIsAlive = Validate-UpdateSource -UpdateSource $UpdateSource
       }
    }

    return $sourceIsAlive
}

Function Validate-UpdateSource() {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [string] $UpdateSource = $NULL
    )

    [bool]$validUpdateSource = $true
    [string]$cabPath = ""

    if ($UpdateSource) {
        $mainRegPath = Get-OfficeCTRRegPath
        $configRegPath = $mainRegPath + "\Configuration"
        $currentplatform = (Get-ItemProperty HKLM:\$configRegPath -Name Platform -ErrorAction SilentlyContinue).Platform
        $updateToVersion = (Get-ItemProperty HKLM:\$configRegPath -Name UpdateToVersion -ErrorAction SilentlyContinue).UpdateToVersion
        $llcc = (Get-ItemProperty HKLM:\$configRegPath -Name ClientCulture -ErrorAction SilentlyContinue).ClientCulture

        $mainCab = "$UpdateSource\Office\Data\v32.cab"
        $bitness = "32"
        if ($currentplatform -eq "x64") {
            $mainCab = "$UpdateSource\Office\Data\v64.cab"
            $bitness = "64"
        }

        if (!($updateToVersion)) {
           $cabXml = Get-CabVersion -FilePath $mainCab
           $updateToVersion = $cabXml.Version.Available.Build
        }

        [xml]$xml = Get-ChannelXml -Bitness $bitness
        $languages = Get-InstalledLanguages

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
               $updateFileExists = Test-Path -Path $fullPath
           }

           if (!($updateFileExists)) {
              $fileExists = $missingFiles.Contains($fullPath)
              if (!($fileExists)) {
                 $missingFiles.Add($fullPath)
                 Write-Host "Source File Missing: $fullPath"
                 Write-Log -Message "Source File Missing: $fullPath" -severity 1 -component "Office 365 Update Anywhere" 
                 <# write log#>
                $lineNum = Get-CurrentLineNumber    
                $filName = Get-CurrentFileName 
                WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Source File Missing: $fullPath"
              }     
              $validUpdateSource = $false
           }
        }

    }
    
    return $validUpdateSource
}

Function Get-InstalledLanguages() {
    [CmdletBinding()]
    Param(
    )
    process {
       $returnLangs = @()
       $mainRegPath = Get-OfficeCTRRegPath

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

       return $returnLangs
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
       <# write log#>
       $lineNum = Get-CurrentLineNumber    
       $filName = Get-CurrentFileName 
       WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Waiting for Update process to Complete..."

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
                                <# write log#>
                                $lineNum = Get-CurrentLineNumber    
                                $filName = Get-CurrentFileName 
                                WriteToLogFile -LNumber $lineNum -FName $filName -ActionError $displayText
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
                                <# write log#>
                                $lineNum = Get-CurrentLineNumber    
                                $filName = Get-CurrentFileName 
                                WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Downloading Update: "
                             }

                             if ($operation.ToUpper().IndexOf("APPLY") -gt -1) {
                                Write-Host "Applying Update: " -NoNewline
                                <# write log#>
                                $lineNum = Get-CurrentLineNumber    
                                $filName = Get-CurrentFileName 
                                WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Applying Update: "
                             }

                             if ($operation.ToUpper().IndexOf("FINALIZE") -gt -1) {
                                Write-Host "Finalizing Update: " -NoNewline
                                <# write log#>
                                $lineNum = Get-CurrentLineNumber    
                                $filName = Get-CurrentFileName 
                                WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Finalizing Update: "
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
              <# write log#>
              $lineNum = Get-CurrentLineNumber    
              $filName = Get-CurrentFileName 
              WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Waiting for Update Timed-Out"
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
       <# write log#>
       $lineNum = Get-CurrentLineNumber    
       $filName = Get-CurrentFileName 
       WriteToLogFile -LNumber $lineNum -FName $filName -ActionError $displayValue

       $totalOperationTime = getOperationTime -OperationStart $totalOperationStart

       if ($updateRunning) {
          if ($failure) {
            Write-Host "Update Failed"
            <# write log#>
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Update Failed"
          } else {
            Write-Host "Update Completed - Total Time: $totalOperationTime"
            <# write log#>
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Update Completed - Total Time: $totalOperationTime"
          }
       } else {
          Write-Host "Update Not Running"
          <# write log#>
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Update Not Running"
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
      <# write log#>
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Invalid UpdateSource. File Not Found: $url"
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
     [bool] $ValidateUpdateSourceFiles = $true
   )

   $newUpdatePath = $UpdatePath
   $newUpdateLong = $UpdatePath

   $detectedChannel = Detect-Channel

   $branchName = $detectedChannel.branch

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
     $pathAlive = Test-UpdateSource -UpdateSource $newUpdatePath -ValidateUpdateSourceFiles $ValidateUpdateSourceFiles
     if (!($pathAlive)) {
        $pathAlive = Test-UpdateSource -UpdateSource $newUpdateLong -ValidateUpdateSourceFiles $ValidateUpdateSourceFiles
        if ($pathAlive) {
           $newUpdatePath = $newUpdateLong
        }
     }
   } catch {
     $pathAlive = $false
   }
   
   if ($pathAlive) {
     return $newUpdatePath
   } else {
     return $UpdatePath
   }
}

function Detect-Channel {
   param( 

   )

   Process {      
      $channelXml = Get-ChannelXml

      $CFGUpdateChannel = (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration -Name UpdateChannel -ErrorAction SilentlyContinue).UpdateChannel
      $CFGOfficeMgmtCOM = (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration -Name OfficeMgmtCOM -ErrorAction SilentlyContinue).OfficeMgmtCOM      
      $UPupdatechannel = (Get-ItemProperty HKLM:\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate -Name UpdateChannel -ErrorAction SilentlyContinue).UpdateChannel      
      $UPupdatepath = (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration -Name updatepath -ErrorAction SilentlyContinue).updatepath
      $officemgmtcom = (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration -Name officemgmtcom -ErrorAction SilentlyContinue).officemgmtcom
      $CFGUpdateUrl = (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration -Name UpdateUrl -ErrorAction SilentlyContinue).UpdateUrl
      $currentBaseUrl = Get-OfficeCDNUrl

      $CurrentChannel = $channelXml.UpdateFiles.baseURL | Where {$_.URL -eq $currentBaseUrl -and $_.branch -notmatch 'Business' }
      
      if($CFGUpdateUrl -ne $null -and $CFGUpdateUrl -like '*officecdn.microsoft.com*'){
        $CurrentChannel = $channelXml.UpdateFiles.baseURL | Where {$_.URL -eq $CFGUpdateUrl -and $_.branch -notmatch 'Business' }  
      }
      if($officemgmtcom -ne $null -and $officemgmtcom -like '*officecdn.microsoft.com*'){
        $CurrentChannel = $channelXml.UpdateFiles.baseURL | Where {$_.URL -eq $officemgmtcom -and $_.branch -notmatch 'Business' }  
      }
      if($UPupdatepath -ne $null -and $UPupdatepath -like '*officecdn.microsoft.com*'){
        $CurrentChannel = $channelXml.UpdateFiles.baseURL | Where {$_.URL -eq $UPupdatepath -and $_.branch -notmatch 'Business' }  
      }
      if($UPupdatechannel -ne $null -and $UPupdatechannel -like '*officecdn.microsoft.com*'){
        $CurrentChannel = $channelXml.UpdateFiles.baseURL | Where {$_.URL -eq $UPupdatechannel -and $_.branch -notmatch 'Business' }  
      }
      if($CFGOfficeMgmtCOM -ne $null -and $CFGOfficeMgmtCOM -like '*officecdn.microsoft.com*'){
        $CurrentChannel = $channelXml.UpdateFiles.baseURL | Where {$_.URL -eq $CFGOfficeMgmtCOM -and $_.branch -notmatch 'Business' }  
      }
      if($CFGUpdateChannel -ne $null -and $CFGUpdateChannel -like '*officecdn.microsoft.com*'){
        $CurrentChannel = $channelXml.UpdateFiles.baseURL | Where {$_.URL -eq $CFGUpdateChannel -and $_.branch -notmatch 'Business' }  
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

       if ($cabPath.ToLower().StartsWith("http")) {
           $webclient = New-Object System.Net.WebClient
           $XMLFilePath = "$env:TEMP/$fileName"
           $XMLDownloadURL= $FilePath
           $webclient.DownloadFile($XMLDownloadURL,$XMLFilePath)
       } else {
         $XMLFilePath = $cabPath
       }

       $tmpName = "VersionDescriptor.xml"
       expand $XMLFilePath $env:TEMP -f:$tmpName | Out-Null
       $tmpName = $env:TEMP + "\VersionDescriptor.xml"
       [xml]$versionXml = Get-Content $tmpName

       return $versionXml
   }
}

function Get-ChannelXml {
   [CmdletBinding()]
   param( 
      [Parameter()]
      [string] $Bitness = "32"
   )

   process {
       $cabPath = "$PSScriptRoot\ofl.cab"

       if (!(Test-Path -Path $cabPath)) {
           $webclient = New-Object System.Net.WebClient
           $XMLFilePath = "$env:TEMP/ofl.cab"
           $XMLDownloadURL = "http://officecdn.microsoft.com/pr/wsus/ofl.cab"
           $webclient.DownloadFile($XMLDownloadURL,$XMLFilePath)
       } else {
           $XMLFilePath = $cabPath
       }

       $tmpName = "o365client_" + $Bitness + "bit.xml"
       expand $XMLFilePath $env:TEMP -f:$tmpName | Out-Null
       $tmpName = $env:TEMP + "\o365client_" + $Bitness + "bit.xml"
       [xml]$channelXml = Get-Content $tmpName

       return $channelXml
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
        [bool] $ValidateUpdateSourceFiles = $true
    )

    Process {
        try {
            $Global:UpdateAnywhereLogPath = $LogPath;
            $Global:UpdateAnywhereLogFileName = $LogName;

            $mainRegPath = Get-OfficeCTRRegPath
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
                               <# write log#>
                                $lineNum = Get-CurrentLineNumber    
                                $filName = Get-CurrentFileName 
                                WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Restoring Saved Update Source $saveUpdateSource"

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
                       <# write log#>
                        $lineNum = Get-CurrentLineNumber    
                        $filName = Get-CurrentFileName 
                        WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "No Update source is set so defaulting to Office CDN"

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
                        <# write log#>
                        $lineNum = Get-CurrentLineNumber    
                        $filName = Get-CurrentFileName 
                        WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Unable to use $currentUpdateSource. Will now use $officeUpdateCDN"

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
               if ($GPOUpdateSource) {
                 $currentUpdateSource = (Get-ItemProperty $GPORegPath -Name updatepath -ErrorAction SilentlyContinue).updatepath
               } else {
                 $currentUpdateSource = (Get-ItemProperty HKLM:\$configRegPath -Name UpdateUrl -ErrorAction SilentlyContinue).UpdateUrl
               }

               if($currentUpdateSource.ToLower().StartsWith("http")){
                   $channelUpdateSource = $currentUpdateSource
               }
               else{
                   $channelUpdateSource = Change-UpdatePathToChannel -UpdatePath $currentUpdateSource -ValidateUpdateSourceFiles $ValidateUpdateSourceFiles
               }

               if ($channelUpdateSource -ne $currentUpdateSource) {
                   if ($GPOUpdateSource) {
                       New-ItemProperty -Path "HKLM:\$GPORegPath2" -Name "updatepath" -Value $channelUpdateSource -PropertyType String -Force -ErrorAction Stop | Out-Null
                   } else {
                       New-ItemProperty -Path "HKLM:\$configRegPath" -Name "UpdateUrl" -Value $channelUpdateSource -PropertyType String -Force -ErrorAction Stop | Out-Null
                   }
                  
                   $channelUpdateSource = $channelUpdateSource
               }

               Write-Host "Starting Update process"
               Write-Host "Update Source: $channelUpdateSource" 
               Write-Log -Message "Will now execute $oc2rcFilePath $oc2rcParams with UpdateSource:$channelUpdateSource" -severity 1 -component "Office 365 Update Anywhere"
               <# write log#>
               $lineNum = Get-CurrentLineNumber    
               $filName = Get-CurrentFileName 
               WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Starting Update process"
               WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Update Source: $channelUpdateSource" 
               WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Will now execute $oc2rcFilePath $oc2rcParams with UpdateSource:$channelUpdateSource"

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
                   Remove-ItemProperty HKLM:\$configRegPath -Name SaveUpdateUrl -Force
               }

            } else {
               $currentUpdateSource = (Get-ItemProperty HKLM:\$configRegPath -Name UpdateUrl -ErrorAction SilentlyContinue).UpdateUrl
               Write-Host "Update Source '$currentUpdateSource' Unavailable"
               Write-Log -Message "Update Source '$currentUpdateSource' Unavailable" -severity 1 -component "Office 365 Update Anywhere"
               <# write log#>
               $lineNum = Get-CurrentLineNumber    
               $filName = Get-CurrentFileName 
               WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Update Source '$currentUpdateSource' Unavailable"
            }

       } catch {
           Write-Log -Message $_.Exception.Message -severity 1 -component $LogFileName
           $fileName = $_.InvocationInfo.ScriptName.Substring($_.InvocationInfo.ScriptName.LastIndexOf("\")+1)
            WriteToLogFile -LNumber $_.InvocationInfo.ScriptLineNumber -FName $fileName -ActionError $_
           throw;
       }
    }
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

Update-Office365Anywhere -WaitForUpdateToFinish $WaitForUpdateToFinish -EnableUpdateAnywhere $EnableUpdateAnywhere -ForceAppShutdown $ForceAppShutdown -UpdatePromptUser $UpdatePromptUser -DisplayLevel $DisplayLevel -UpdateToVersion $UpdateToVersion -LogPath $LogPath -LogName $LogName -ValidateUpdateSourceFiles $ValidateUpdateSourceFiles



