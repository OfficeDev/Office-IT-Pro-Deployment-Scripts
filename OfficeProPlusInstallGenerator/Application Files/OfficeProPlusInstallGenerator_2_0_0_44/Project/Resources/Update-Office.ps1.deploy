[CmdletBinding()]
Param(   

    [Parameter()]
    [string] $updatetoversion = $null,
    
    [Parameter()]
    [string] $channel = $null,

    [Parameter()]
    [bool] $DisplayLevel = $false

)

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

Function Get-LCID(){
    [CmdletBinding()]
        Param(
        [Parameter(Mandatory=$true)]
        [string] $UpdateSource = $NULL
    )

    if ($UpdateSource) {
        $mainRegPath = Get-OfficeCTRRegPath
        $configRegPath = $mainRegPath + "\Configuration"
        $llcc = (Get-ItemProperty HKLM:\$configRegPath -Name ClientCulture -ErrorAction SilentlyContinue).ClientCulture
        $cabversion = (Get-Item ($UpdateSource + "\Office\Data\" + "*\") | Where-Object {$_.Mode -eq "d-----"})
        $cabverdir = $cabversion.Name
    }    

       switch ($llcc)
        {        ar-sa{$lcid = "1025"}        bg-bg{$lcid = "1026"}        zh-cn{$lcid = "2052"}        zh-tw{$lcid = "1028"}        hr-hr{$lcid = "1050"}        cs-cz{$lcid = "1029"}        da-dk{$lcid = "1030"}        nl-nl{$lcid = "1043"}        en-us{$lcid = "1033"}        et-ee{$lcid = "1061"}        fi-fi{$lcid = "1035"}        fr-fr{$lcid = "1036"}        de-de{$lcid = "1031"}        el-gr{$lcid = "1032"}        he-il{$lcid = "1037"}        hi-in{$lcid = "1081"}        hu-hu{$lcid = "1038"}        id-id{$lcid = "1057"}        it-it{$lcid = "1040"}        ja-jp{$lcid = "1041"}        kk-kz{$lcid = "1087"}        ko-kr{$lcid = "1042"}        lv-lv{$lcid = "1062"}        lt-lt{$lcid = "1063"}        ms-my{$lcid = "1086"}        nb-no{$lcid = "1044"}        pl-pl{$lcid = "1045"}        pt-br{$lcid = "1046"}        pt-pt{$lcid = "2070"}        ro-ro{$lcid = "1048"}        ru-ru{$lcid = "1049"}        sr-latn-cs{$lcid = "2074"}        sk-sk{$lcid = "1051"}        sl-si{$lcid = "1060"}        es-es{$lcid = "3082"}        sv-se{$lcid = "1053"}        th-th{$lcid = "1054"}        tr-tr{$lcid = "1055"}        uk-ua{$lcid = "1058"}        vi-vn{$lcid = "1066"}
        }

       Return $lcid        
    }

Function Test-UpdateSource() {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [string] $UpdateSource = $NULL
    )

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

    return $sourceIsAlive
}

Function Validate-UpdateSource() {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [string] $UpdateSource = $NULL
    )

    [bool]$validUpdateSource = $false
    [string]$cabPath = ""

    if ($UpdateSource) {
        $mainRegPath = Get-OfficeCTRRegPath
        $configRegPath = $mainRegPath + "\Configuration"
        $currentplatform = (Get-ItemProperty HKLM:\$configRegPath -Name Platform -ErrorAction SilentlyContinue).Platform
        $updateToVersion = (Get-ItemProperty HKLM:\$configRegPath -Name UpdateToVersion -ErrorAction SilentlyContinue).UpdateToVersion
        $llcc = (Get-ItemProperty HKLM:\$configRegPath -Name ClientCulture -ErrorAction SilentlyContinue).ClientCulture
        $cabversion = (Get-Item ($UpdateSource + "\Office\Data\" + "*\") | Where-Object {$_.Mode -eq "d-----"})
        $cabverdir = $cabversion.Name
        $runcid = Get-LCID -UpdateSource $UpdateSource
        $setlcid = $lcid

        
        if ($updateToVersion) {
            if ($currentplatform.ToLower() -eq "x86") {
               $cabPath = $UpdateSource + "\Office\Data\v32_" + $updateToVersion + ".cab"
            }
            if ($currentplatform.ToLower() -eq "x64") {
               $cabPath = $UpdateSource + "\Office\Data\v64_" + $updateToVersion + ".cab"
            }
        } 
        if ($runcid) {
            if ($currentplatform.ToLower() -eq "X86"){
                $cabpathi = $UpdateSource + "\Office\Data\" + $cabverdir + "\i86" + $runcid + ".cab"
                $cabpaths = $UpdateSource + "\Office\Data\" + $cabverdir + "\s86" + $runcid + ".cab"
            }
            else{
                $cabpathi = $UpdateSource + "\Office\Data\" + $cabverdir + "\i64" + $runcid + ".cab"
                $cabpaths = $UpdateSource + "\Office\Data\" + $cabverdir + "\s64" + $runcid + ".cab"
            }
              
            
        }
        
        if($UpdateSource.ToLower().StartsWith("http")){        
            if ($currentplatform.ToLower() -eq "x86") {
               $cabPath = $UpdateSource + "\Office\Data\v32.cab"
            }
            else {
               $cabPath = $UpdateSource + "\Office\Data\v64.cab"
            }            
        }
        else{
            if ($currentplatform.ToLower() -eq "x86") {
               $cabPath = $UpdateSource + "\Office\Data\v32.cab"
            }
            else {
               $cabPath = $UpdateSource + "\Office\Data\v64.cab"
            }
        }

        if ($cabPath.ToLower().StartsWith("http")) {
           $cabPath = $cabPath.Replace("\", "/")
           $validUpdateSource = Test-URL -url $cabPath
        } else {
           $validUpdateSource = Test-Path -Path $cabPath
           $validUpdateSourcei = Test-Path -Path $cabpathi
           $validUpdateSources = Test-Path -Path $cabpaths
        }


        
        if (!$validUpdateSource) {
           Write-Host "Invalid UpdateSource. File Not Found: $cabPath"
        }
        if (!$validUpdateSourcei) {
            Write-Host "Invalid UpdateSource. File Not Found: $cabpathi"
        }
        if (!$validUpdateSources){
            Write-Host "Invalid Updateource. File Not Found: $cabpaths"
        }
    }
    
        return $validUpdateSource
}



Function Update-Office() {
<#


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
        [string] $Channel = $NULL
        
    )

    Process {
        try {
            $Global:UpdateAnywhereLogPath = $LogPath;
            $Global:UpdateAnywhereLogFileName = $LogName;

            $mainRegPath = Get-OfficeCTRRegPath
            $configRegPath = $mainRegPath + "\Configuration"
            

            
            

            
             $currentUpdateSource = (Get-ItemProperty HKLM:\$configRegPath -Name UpdateUrl -ErrorAction SilentlyContinue).UpdateUrl
            
            
                       
            $clientFolder = (Get-ItemProperty HKLM:\$configRegPath -Name ClientFolder -ErrorAction SilentlyContinue).ClientFolder

            $CDNUpdatePath = (Get-ItemProperty HKLM:\$configRegPath -Name CDNBaseURL -ErrorAction SilentlyContinue).CDNBaseURL

            if($Channel.ToLower().StartsWith("deferred")){
                $CDNBasePath = "http://officecdn.microsoft.com/pr/7ffbc6bf-bc32-4f92-8982-f9dd17fd3114"
            }
            if($Channel.ToLower().StartsWith("firstreleasedeferred")){
                $CDNBasePath = "http://officecdn.microsoft.com/pr/b8f9b850-328d-4355-9145-c59439a0c4cf"
            }
            if($Channel.ToLower().StartsWith("current")){
                $CDNBasePath = "http://officecdn.microsoft.com/pr/492350f6-3a01-4f97-b9c0-c7c6ddf67d60"
            }
            if($Channel.ToLower().StartsWith("firstreleasecurrent")){
                $CDNBasePath = "http://officecdn.microsoft.com/pr/64256afe-f5d9-4f86-8936-8840a6a4f5be"
            }
            
            if($CDNUpdatePath){
                Set-Reg -Hive "HKLM" -keyPath $configRegPath -ValueName "CDNBaseURL" -Value $CDNBasePath -Type String
             }

            
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
			$oc2rcParams += " updatepromptuser=false"
    
            $UpdateSource = "http"
            if ($currentUpdateSource) {
              If ($currentUpdateSource.StartsWith("\\",1)) {
                 $UpdateSource = "UNC"
              }
            }
            else{
            Set-Reg -Hive "HKLM" -keyPath $configRegPath -ValueName "CDNBaseURL" -Value $CDNBasePath -Type String
            }
            
            Restart-Service ClickToRunSvc         

           if ($currentUpdateSource) {
               $currentUpdateSource = (Get-ItemProperty HKLM:\$configRegPath -Name UpdateUrl -ErrorAction SilentlyContinue).UpdateUrl
               if($currentUpdateSource.ToLower().StartsWith("http")){
                   $channelUpdateSource = $currentUpdateSource
               }
               else{
                   $channelUpdateSource = Change-UpdatePathToChannel -UpdatePath $currentUpdateSource -Channel $Channel                   
               }

               if ($channelUpdateSource -ne $currentUpdateSource) {                   
                     Set-Reg -Hive "HKLM" -keyPath $configRegPath -ValueName "UpdateUrl" -Value $channelUpdateSource -Type String                  
               }
               #$isPathGood = Test-UpdateSource -UpdateSource $channelUpdateSource

               #if(!$isPathGood){
               #     $pathToErase = 'HKLM:\'+ $configRegPath
               #     Remove-ItemProperty -Path $pathToErase -Name UpdateURL
               #}

               Write-Host "Starting Update process"
               Write-Host "Update Source: $currentUpdateSource" 
               Write-Log -Message "Will now execute $oc2rcFilePath $oc2rcParams with UpdateSource:$currentUpdateSource" -severity 1 -component "Office 365 Update Anywhere"
               StartProcess -execFilePath $oc2rcFilePath -execParams $oc2rcParams

               if ($WaitForUpdateToFinish) {
                    Wait-ForOfficeCTRUpadate
               }

           } else {
               Write-Host "Starting Update process"
               Write-Host "Update Source: $currentUpdateSource" 
               Write-Log -Message "Will now execute $oc2rcFilePath $oc2rcParams with UpdateSource:$currentUpdateSource" -severity 1 -component "Office 365 Update Anywhere"
               StartProcess -execFilePath $oc2rcFilePath -execParams $oc2rcParams

               if ($WaitForUpdateToFinish) {
                    $outputText= Wait-ForOfficeCTRUpadate
               }
           }

       } catch {
           Write-Log -Message $_.Exception.Message -severity 1 -component $LogFileName
           throw;
       }
       return $outputText
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
        $returnVar = ""
    }
    
    process {
    
       Write-Host "Waiting for Update process to Complete..."
       $returnVar+= "Waiting for Update process to Complete..."

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
                                $returnVar += $displayText
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
            $returnVar+= "Update Failed"
          } else {
            Write-Host "Update Completed - Total Time: $totalOperationTime"
            $returnVar+= "Update Completed - Total Time: $totalOperationTime"
          }
       } else {
          Write-Host "Update Not Running"
          $returnVar+= "Update Not Running"
       } 
       return $returnVar
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
     [string] $Channel
   )

   $newUpdatePath = $UpdatePath

   $detectedChannel = $Channel

   $branchName = $detectedChannel

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
   
   if ($madeChange) {
     return $newUpdatePath
   } else {
     return $UpdatePath
   }
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

function Get-ChannelXml {
   [CmdletBinding()]
   param( 
      
   )

   process {
       $cabPath = "$PSScriptRoot\ofl.cab"

       if (!(Test-Path -Path $cabPath)) {
           $webclient = New-Object System.Net.WebClient
           $XMLFilePath = "$env:TEMP/ofl.cab"
           $XMLDownloadURL = "http://officecdn.microsoft.com/pr/wsus/ofl.cab"
           $webclient.DownloadFile($XMLDownloadURL,$XMLFilePath)
       }

       $tmpName = "o365client_64bit.xml"
       expand $XMLFilePath $env:TEMP -f:$tmpName | Out-Null
       $tmpName = $env:TEMP + "\o365client_64bit.xml"
       [xml]$channelXml = Get-Content $tmpName

       return $channelXml
   }

}

Update-Office -DisplayLevel $DisplayLevel -UpdateToVersion $updatetoversion -Channel $channel



