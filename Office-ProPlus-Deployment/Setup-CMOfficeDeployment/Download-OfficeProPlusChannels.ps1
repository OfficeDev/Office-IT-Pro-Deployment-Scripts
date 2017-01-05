try {
$enumDef = "
using System;
       [FlagsAttribute]
       public enum Bitness
       {
          Both = 0,
          v32 = 1,
          v64 = 2
       }
"
Add-Type -TypeDefinition $enumDef -ErrorAction SilentlyContinue
} catch { }

try {
$enumDef = "
using System;
       [FlagsAttribute]
       public enum OfficeBranch
       {
          FirstReleaseCurrent = 0,
          Current = 1,
          FirstReleaseBusiness = 2,
          Business = 3,
          CMValidation = 4
       }
"
Add-Type -TypeDefinition $enumDef -ErrorAction SilentlyContinue
} catch { }

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

function Download-OfficeProPlusChannels{
<#
.SYNOPSIS
Downloads each Office ProPlus Channel with installation files
.DESCRIPTION
This script will dynamically downloaded the most current Office ProPlus version for each deployment Channel
.PARAMETER Version
The version number you wish to download. For example: 16.0.6228.1010
.PARAMETER TargetDirectory
Required. Where all the channels will be downloaded. Each channel then goes into a folder of the same name as the channel.
.PARAMETER Languages
Array of Microsoft language codes. Will throw error if provided values don't match the validation set. Defaults to "en-us"
("en-us","ar-sa","bg-bg","zh-cn","zh-tw","hr-hr","cs-cz","da-dk","nl-nl","et-ee","fi-fi","fr-fr","de-de","el-gr","he-il","hi-in","hu-hu","id-id","it-it",
"ja-jp","kk-kz","ko-kr","lv-lv","lt-lt","ms-my","nb-no","pl-pl","pt-br","pt-pt","ro-ro","ru-ru","sr-latn-rs","sk-sk","sl-si","es-es","sv-se","th-th",
"tr-tr","uk-ua")
.PARAMETER Bitness
v32, v64, or Both. What bitness of office you wish to download. Defaults to Both.
.PARAMETER OverWrite
If this parameter is specified then existing files will be overwritten.
.PARAMETER Branches
An array of the Branches you wish to download (This parameter is left for legacy usage)
.PARAMETER Channels
An array of the Channels you wish to download. Defaults to all available channels except First Release Current
.PARAMETER NumVersionsToKeep
This parameter controls the number of versions to retain. Any older versions will be deleted.
.PARAMETER UseChannelFolderShortName
This parameter change the folder name that the scripts creates for each Channel folder. For example if this paramter is set to $true then the Current Channel folder will be named "CC"
.PARAMETER NumOfRetries
This parameter Controls the number of times the script will retry if a failure happens
.PARAMETER IncludeChannelInfo
This parameter Controls whether the ofl.cab file is downloaded and cached in the root of the TargetDirectory folder
.PARAMETER DownloadPreviousVersionIfThrottled
This parameter will force the function to download the previous version if the current version is still being throttled
.Example
Download-OfficeProPlusChannels -TargetDirectory "\\server\updateshare"
Default downloads all available channels of the most recent version for both bitnesses into an update source. Downloads the English language pack by default if language is not specified.
.Link
https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts
#>

Param(
    [Parameter()]
    [string] $Version,

    [Parameter(Mandatory=$true)]
    [string] $TargetDirectory,

    [Parameter()]
    [ValidateSet("en-us","ar-sa","bg-bg","zh-cn","zh-tw","hr-hr","cs-cz","da-dk","nl-nl","et-ee","fi-fi","fr-fr","de-de","el-gr","he-il","hi-in","hu-hu","id-id","it-it",
                "ja-jp","kk-kz","ko-kr","lv-lv","lt-lt","ms-my","nb-no","pl-pl","pt-br","pt-pt","ro-ro","ru-ru","sr-latn-rs","sk-sk","sl-si","es-es","sv-se","th-th",
                "tr-tr","uk-ua","vi-vn")]
    [string[]] $Languages = ("en-us"),

    [Parameter()]
    [ValidateSet("af-za","sq-al","am-et","hy-am","as-in","az-latn-az","eu-es","be-by","bn-bd","bn-in","bs-latn-ba","ca-es","prs-af","fil-ph","gl-es","ka-ge","gu-in","is-is","ga-ie","kn-in",
                "km-kh","sw-ke","kok-in","ky-kg","lb-lu","mk-mk","ml-in","mt-mt","mi-nz","mr-in","mn-mn","ne-np","nn-no","or-in","fa-ir","pa-in","quz-pe","gd-gb","sr-cyrl-rs","sr-cyrl-ba",
                "sd-arab-pk","si-lk","ta-in","tt-ru","te-in","tk-tm","ur-pk","ug-cn","uz-latn-uz","ca-es-valencia","cy-gb")]
    [string[]] $PartialLanguages,

    [Parameter()]
    [ValidateSet("ha-latn-ng","ig-ng","xh-za","zu-za","rw-rw","ps-af","rm-ch","nso-za","tn-za","wo-sn","yo-ng")]
    [string[]] $ProofingLanguages,

    [Parameter()]
    [Bitness] $Bitness = 0,

    [Parameter()]
    [int] $NumVersionsToKeep = 2,

    [Parameter()]
    [bool] $UseChannelFolderShortName = $true,

    [Parameter()]
    [bool] $OverWrite = $false,

    [Parameter()]
    [OfficeBranch[]] $Branches,

    [Parameter()]
    [OfficeChannel[]] $Channels = (0, 1, 2, 3),

    [Parameter()]
    [int] $NumOfRetries = 5,

    [Parameter()]
    [bool] $IncludeChannelInfo = $false,

    [Parameter()]
    [bool] $DownloadPreviousVersionIfThrottled = $false
)

#create array for all languages including core, partial, and proofing
$allLanguages = @();
$allLanguages += , $Languages
$allLanguages += , $PartialLanguages
$allLanguages += , $ProofingLanguages


$BranchesOrChannels = @()

if($Branches.Count -gt 0)
{
    foreach ($branchName in $Branches) {
      $channelConvertName = ConvertBranchNameToChannelName -BranchName $branchName
      $BranchesOrChannels += $channelConvertName
    }
}
else{
    $BranchesOrChannels = $Channels
}
      
$numberOfFiles = (($BranchesOrChannels.Count) * ((($allLanguages.Count + 1)*3) + 5))

[bool]$downloadSuccess = $TRUE;
For($i=1; $i -le $NumOfRetries; $i++){#loops through download process in the event of a failure in order to retry

    try{
        $XMLFilePath = "$env:TEMP/ofl.cab"
        $XMLDownloadURL = "http://officecdn.microsoft.com/pr/wsus/ofl.cab"

        DownloadFile -url $XMLDownloadURL -targetFile $XMLFilePath

        if ($IncludeChannelInfo) {
            Copy-Item -Path $XMLFilePath -Destination "$TargetDirectory\ofl.cab"
        }

        if($Bitness -eq [Bitness]::Both -or $Bitness -eq [Bitness]::v32){
            $32XMLFileName = "o365client_32bit.xml"
            expand $XMLFilePath $env:TEMP -f:$32XMLFileName | Out-Null
            $32XMLFilePath = $env:TEMP + "\o365client_32bit.xml"
            [xml]$32XML = Get-Content $32XMLFilePath
            $xmlArray = ($32XML)
        }

        if($Bitness -eq [Bitness]::Both -or $Bitness -eq [Bitness]::v64){
            $64XMLFileName = "o365client_64bit.xml"
            expand $XMLFilePath $env:TEMP -f:$64XMLFileName | Out-Null
            $64XMLFilePath = $env:TEMP + "\o365client_64bit.xml"
            [xml]$64XML = Get-Content $64XMLFilePath
            if($xmlArray -ne $null){
                $xmlArray = ($32XML,$64XML)
                $numberOfFiles = $numberOfFiles * 2
            }else{
                $xmlArray = ($64XML)
            }
        }

        $j = 0
        $b = 0
        $BranchCount = $BranchesOrChannels.Count * 2

        #loop to download files
        $xmlArray | %{
            $CurrentVersionXML = $_
    
            $currentBitness = "32-Bit"
            if ($CurrentVersionXML.OuterXml.Contains("Architecture: 64 Bit")) {
                $currentBitness = "64-Bit"
            }

            Write-Host
            Write-Host "Downloading Bitness : $currentBitness"
            <# write log#>
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Downloading Bitness : $currentBitness"

            #loop for each branch
            $BranchesOrChannels | %{
                $currentBranch = $_
                $b++

                $Version = ""
                $PreviousVersion = ""
                $NewestVersion = ""
                $Throttle = ""
                $VersionFile = ""

                Write-Progress -id 1 -Activity "Downloading Channel" -status "Channel: $($currentBranch.ToString()) : $currentBitness" -percentComplete ($b / $BranchCount *100) 

                <# write log#>
                $lineNum = Get-CurrentLineNumber    
                $filName = Get-CurrentFileName 
                WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Downloading Channel: $currentBranch"

                $FolderName = $($_.ToString())

                if ($UseChannelFolderShortName) {
                   $FolderName = ConvertChannelNameToShortName -ChannelName $FolderName  
                }
       
                $baseURL = $CurrentVersionXML.UpdateFiles.baseURL | ? branch -eq $_.ToString() | %{$_.URL};
                if(!(Test-Path "$TargetDirectory\$FolderName\")){
                    New-Item -Path "$TargetDirectory\$FolderName\" -ItemType directory -Force | Out-Null
                }
                if(!(Test-Path "$TargetDirectory\$FolderName\Office")){
                    New-Item -Path "$TargetDirectory\$FolderName\Office" -ItemType directory -Force | Out-Null
                }
                if(!(Test-Path "$TargetDirectory\$FolderName\Office\Data")){
                    New-Item -Path "$TargetDirectory\$FolderName\Office\Data" -ItemType directory -Force | Out-Null
                }

                if([String]::IsNullOrWhiteSpace($Version) -or [String]::IsNullOrWhiteSpace($Throttle)){
                    $versionReturn = GetVersionBasedOnThrottle -Channel $currentBranch -Version $Version -currentVerXML $CurrentVersionXML
                    if ($DownloadPreviousVersionIfThrottled) {
                        if ($versionReturn.Throttle -ge 1000) {
                          $Version = $versionReturn.NewestVersion
                        } else {
                          $Version = $versionReturn.PreviousVesion
                        }
                    }
                    $NewestVersion = $versionReturn.NewestVersion
                    $PreviousVersion = $versionReturn.PreviousVesion
                    $Throttle = $versionReturn.Throttle
                }             

                if([String]::IsNullOrWhiteSpace($Version)){
                    #get base .cab to get current version
                    $baseCabFile = $CurrentVersionXML.UpdateFiles.File | ? rename -ne $null
                    $url = "$baseURL$($baseCabFile.relativePath)$($baseCabFile.rename)"
                    $destination = "$TargetDirectory\$FolderName\Office\Data\$($baseCabFile.rename)"

                    DownloadFile -url $url -targetFile $destination

                    expand $destination $env:TEMP -f:"VersionDescriptor.xml" | Out-Null
                    $baseCabFileName = $env:TEMP + "\VersionDescriptor.xml"
                    [xml]$vdxml = Get-Content $baseCabFileName
                    $currentVersion = $vdxml.Version.Available.Build;
                    Remove-Item -Path $baseCabFileName
                }else{
                    $currentVersion = $Version

                    $relativePath = $_.relativePath -replace "`%version`%", $currentVersion
                    $fileName = "/Office/Data/v32_$currentVersion.cab"
                    $url = "$baseURL$relativePath$fileName"

                    try {
                        Invoke-WebRequest -Uri $url -ErrorAction Stop | Out-Null
                    } catch {
                      Write-Host "`t`tVersion Not Found: $currentVersion"
                      <# write log#>
                        $lineNum = Get-CurrentLineNumber    
                        $filName = Get-CurrentFileName 
                        WriteToLogFile -LNumber $lineNum -FName $fileName -ActionError "Version Not Found: $currentVersion"
                      return 
                    }
                }

                if ($currentBitness.Contains("32")) {
                   $VersionFile = "$TargetDirectory\$FolderName\Office\Data\v32_$currentVersion.cab"
                } else {
                   $VersionFile = "$TargetDirectory\$FolderName\Office\Data\v64_$currentVersion.cab"
                }
                
                if (($Throttle -lt 1000) -and ($DownloadPreviousVersionIfThrottled)) {
                   Write-Host "`tDownloading Channel: $currentBranch - Version: $currentVersion (Using previous version instead of Thottled Version: $NewestVersion - Thottle: $Throttle)"
                } else {
                   Write-Host "`tDownloading Channel: $currentBranch - Version: $currentVersion"
                }


                if(!(Test-Path "$TargetDirectory\$FolderName\Office\Data\$currentVersion")){
                    New-Item -Path "$TargetDirectory\$FolderName\Office\Data\$currentVersion" -ItemType directory -Force | Out-Null
                }
				if(!(Test-Path "$TargetDirectory\$FolderName\Office\Data\$currentVersion\Experiment")){
                    New-Item -Path "$TargetDirectory\$FolderName\Office\Data\$currentVersion\Experiment" -ItemType directory -Force | Out-Null
 				}
                if(!(Test-Path "$TargetDirectory\$FolderName\Office\Data\$currentVersion\Experiment")){
                    New-Item -Path "$TargetDirectory\$FolderName\Office\Data\$currentVersion\Experiment" -ItemType directory -Force | Out-Null
                }

                $numberOfFiles = 0
                $j = 0

                $CurrentVersionXML.UpdateFiles.File | ? language -eq "0" | 
                %{
                   $numberOfFiles ++
                }

                $allLanguages | 
                %{
                    #LANGUAGE LOGIC HERE
                    $languageId  = [globalization.cultureinfo]::GetCultures("allCultures") | ? Name -eq $_ | %{$_.LCID}
                    $CurrentVersionXML.UpdateFiles.File | ? language -eq $languageId | 
                            %{
                   $numberOfFiles ++
                }
                }


                #basic files
                $CurrentVersionXML.UpdateFiles.File | ? language -eq "0" | 
                %{
                    $name = $_.name -replace "`%version`%", $currentVersion
                    $relativePath = $_.relativePath -replace "`%version`%", $currentVersion
                    $url = "$baseURL$relativePath$name"
                    $fileType = $name.split('.')[$name.split['.'].Count - 1]
                    $bitnessValue = $currentBitness.split('-')[0].ToString()
                    $destination = "$TargetDirectory\$FolderName$relativePath$name"
                               
                    for( $retryCount = 0; $retryCount -lt 3;  $retryCount++) {
                        try {
                            $hashFileName = $name.replace("dat","hash")
                            $cabFile = "$TargetDirectory\$FolderName"+$relativePath.Replace('/','\')+"s"+$bitnessValue+"0.cab"
                            $noneHashLocation = $CurrentVersionXML.UpdateFiles.File | ? name -eq $name |%{$_.hashLocation}  

                            $downloadfile = $true
                            $pathTest = Test-Path -Path $destination

                            if ($OverWrite) {
                              $pathTest = $false
                            }

                            if ($pathTest) {
                               if($fileType -eq 'dat')
                               {                                   
                                   $hashMatch = Check-FileHash -FilePath $destination -CabFile $cabFile
                                   if($hashMatch)
                                   {
                                      $downloadfile = $false
                                   }           
                               } else {
                                 $downloadfile = $false
                               }
                            }

                            if ($downloadfile) { 
                               DownloadFile -url $url -targetFile $destination
                     
                               if($fileType -eq 'dat')
                               {                                   
                                   $hashMatch = Check-FileHash -FilePath $destination -CabFile $cabFile
                                   if(!($hashMatch))
                                   {
                                      throw "$name file hash is not correct";
                                   }           
                               }
                           }

                           break;
                        }
                        catch{
                            $OverWrite = $true 
                            if ($retryCount -eq 2) {
                                throw 
                            }        
                        }
                    }

                    $j = $j + 1

                    if (($Throttle -lt 1000) -and ($DownloadPreviousVersionIfThrottled)) {
                       Write-Progress -id 2 -ParentId 1 -Activity "Downloading Channel Files" -status "Channel: $($currentBranch.ToString()) - Version: $currentVersion (Using previous version instead of Thottled Version: $NewestVersion - Thottle: $Throttle)" -percentComplete ($j / $numberOfFiles *100)
                    } else {
                       Write-Progress -id 2 -ParentId 1 -Activity "Downloading Channel Files" -status "Channel: $($currentBranch.ToString()) - Version: $currentVersion" -percentComplete ($j / $numberOfFiles *100)
                    }
                }

                #language files
                $allLanguages | 
                %{
                    #LANGUAGE LOGIC HERE
                    $languageId  = [globalization.cultureinfo]::GetCultures("allCultures") | ? Name -eq $_ | %{$_.LCID}
					$bitnessValue = $currentBitness.split('-')[0].ToString()
                    $CurrentVersionXML.UpdateFiles.File | ? language -eq $languageId | 

                    %{
                    
                    $name = $_.name -replace "`%version`%", $currentVersion                    
                    for( $retryCount = 0; $retryCount -lt 3;  $retryCount++) {
                            try {
                        
                                $fileType = $name.split('.')[$name.split['.'].Count - 1]
                                $relativePath = $_.relativePath -replace "`%version`%", $currentVersion
                                $url = "$baseURL$relativePath$name"
                                $destination = "$TargetDirectory\$FolderName"+$relativePath.replace('/','\')+"$name"

                                $cabFile = "$TargetDirectory\$FolderName"+$relativePath.Replace('/','\')+"s"+$bitnessValue+$languageId+".cab"
                     
                                $downloadfile = $true
                                $pathTest = Test-Path -Path $destination

                                if ($OverWrite) {
                                  $pathTest = $false
                                }

                                if ($pathTest) {
                                   if($fileType -eq 'dat')
                                   {                                   
                                       $hashMatch = Check-FileHash -FilePath $destination -CabFile $cabFile
                                       if($hashMatch)
                                       {
                                          $downloadfile = $false
                                       }           
                                   } else {
                                     $downloadfile = $false
                                   }
                                }

                                if ($downloadfile) {
                                   DownloadFile -url $url -targetFile $destination

                                   if($fileType -eq 'dat')
                                   {                                   
                                       $hashMatch = Check-FileHash -FilePath $destination -CabFile $cabFile
                                       if(!($hashMatch))
                                       {
                                          throw "$name file hash is not correct"
                                       }           
                                   }
                                }
                                    
                                break;
                            } catch {
                                $OverWrite = $true 
                                if ($retryCount -eq 2) {
                                   throw 
                                }
                            }
                        }

                        $j = $j + 1
                        Write-Progress -id 2 -ParentId 1 -Activity "Downloading Channel Files" -status "Channel: $($currentBranch.ToString())" -percentComplete ($j / $numberOfFiles *100)
                    }
                }

                #Copy Version file and overwrite the v32.cab or v64.cab file
                if (Test-Path -Path $VersionFile) {
                   $parentPath = Split-Path -parent $VersionFile

                   if ($currentBitness.Contains("32")) {
                     Copy-Item -Path $VersionFile -Destination "v32.cab" -Force | Out-Null
                   } else {
                     Copy-Item -Path $VersionFile -Destination "v64.cab" -Force | Out-Null
                   }
                }
            }

        }

    }
    catch 
    {
        #if download fails, displays error, continues loop
        $errorMessage = $computer + ": " + $_
        Write-Host $errorMessage -ForegroundColor White -BackgroundColor Red
        $downloadSuccess = $FALSE;
        $fileName = $_.InvocationInfo.ScriptName.Substring($_.InvocationInfo.ScriptName.LastIndexOf("\")+1)
        WriteToLogFile -LNumber $_.InvocationInfo.ScriptLineNumber -FName $fileName -ActionError $_
    }

    if($downloadSuccess){#if download succeeds, breaks out of loop
        break
    }

}#end of for loop

Write-Host
PurgeOlderVersions $TargetDirectory $NumVersionsToKeep $BranchesOrChannels

}

function Get-OfficeProPlusChannelInfo {
<#
.SYNOPSIS
Downloads each Office ProPlus Channel with installation files
.DESCRIPTION
This script will display the latest version from each Channel with the throttle value of the latest version and the previous version from the latest
.PARAMETER Channels
An array of the Channels you wish to display information
.Example
Get-OfficeProPlusChannelInfo -Channels Current
This will only display the information for the Current Channel
.Link
https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts
#>

Param(
    [Parameter()]
    [OfficeChannel[]] $Channels = (0, 1, 2, 3)
)

begin {
    $defaultDisplaySet = 'Channel', 'LatestVersion', 'Throttle', 'PreviousVersion'

    $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet(‘DefaultDisplayPropertySet’,[string[]]$defaultDisplaySet)
    $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
}

Process {
    $results = new-object PSObject[] 0;

    $BranchesOrChannels = $Channels
      
    try{
        $XMLFilePath = "$env:TEMP/ofl.cab"
        $XMLDownloadURL = "http://officecdn.microsoft.com/pr/wsus/ofl.cab"

        DownloadFile -url $XMLDownloadURL -targetFile $XMLFilePath

        if ($IncludeChannelInfo) {
            Copy-Item -Path $XMLFilePath -Destination "$TargetDirectory\ofl.cab"
        }

        $Bitness = [Bitness]::v32

        if($Bitness -eq [Bitness]::Both -or $Bitness -eq [Bitness]::v32){
            $32XMLFileName = "o365client_32bit.xml"
            expand $XMLFilePath $env:TEMP -f:$32XMLFileName | Out-Null
            $32XMLFilePath = $env:TEMP + "\o365client_32bit.xml"
            [xml]$32XML = Get-Content $32XMLFilePath
            $xmlArray = ($32XML)
        }

        if($Bitness -eq [Bitness]::Both -or $Bitness -eq [Bitness]::v64){
            $64XMLFileName = "o365client_64bit.xml"
            expand $XMLFilePath $env:TEMP -f:$64XMLFileName | Out-Null
            $64XMLFilePath = $env:TEMP + "\o365client_64bit.xml"
            [xml]$64XML = Get-Content $64XMLFilePath
            if($xmlArray -ne $null){
                $xmlArray = ($32XML,$64XML)
                $numberOfFiles = $numberOfFiles * 2
            }else{
                $xmlArray = ($64XML)
            }
        }

        $xmlArray | %{
            $CurrentVersionXML = $_
    
            $currentBitness = "32-Bit"
            if ($CurrentVersionXML.OuterXml.Contains("Architecture: 64 Bit")) {
                $currentBitness = "64-Bit"
            }

            #loop for each branch
            $BranchesOrChannels | %{
                $currentBranch = $_

                $Version = ""
                $PreviousVersion = ""
                $NewestVersion = ""
                $Throttle = ""
                $VersionFile = ""

                $FolderName = $($_.ToString())
       
                $baseURL = $CurrentVersionXML.UpdateFiles.baseURL | ? branch -eq $_.ToString() | %{$_.URL};

                $versionReturn = GetVersionBasedOnThrottle -Channel $currentBranch -Version $Version -currentVerXML $CurrentVersionXML
                if ($DownloadPreviousVersionIfThrottled) {
                    if ($versionReturn.Throttle -ge 1000) {
                        $Version = $versionReturn.NewestVersion
                    } else {
                        $Version = $versionReturn.PreviousVesion
                    }
                }
                $NewestVersion = $versionReturn.NewestVersion
                $PreviousVersion = $versionReturn.PreviousVersion
                $Throttle = $versionReturn.Throttle       

                $TargetDirectory = $env:TEMP
                $destFolderPath = "$TargetDirectory\$FolderName\Office\Data"
                [system.io.directory]::CreateDirectory($destFolderPath) | Out-Null

                $object = New-Object PSObject -Property @{Channel = $currentBranch; 
                                                          LatestVersion = $NewestVersion; 
                                                          Throttle = $Throttle;
                                                          PreviousVersion = $PreviousVersion }
                $object | Add-Member MemberSet PSStandardMembers $PSStandardMembers
                $results += $object
            }

        }

    }
    catch 
    {
        $errorMessage = $computer + ": " + $_
        Write-Host $errorMessage -ForegroundColor White -BackgroundColor Red
    }

    if($downloadSuccess){
        #if download succeeds, breaks out of loop
        break
    }

    return $results
}
}

function DownloadFile($url, $targetFile) {

  for($t=1;$t -lt 10; $t++) {
   try {
       $uri = New-Object "System.Uri" "$url"
       $request = [System.Net.HttpWebRequest]::Create($uri)
       $request.set_Timeout(15000) #15 second timeout

       $response = $request.GetResponse()
       $totalLength = [System.Math]::Floor($response.get_ContentLength()/1024)
       $responseStream = $response.GetResponseStream()
       $targetStream = New-Object -TypeName System.IO.FileStream -ArgumentList $targetFile.replace('/','\'), Create
       $buffer = new-object byte[] 8192KB
       $count = $responseStream.Read($buffer,0,$buffer.length)
       $downloadedBytes = $count

       while ($count -gt 0)
       {
           $targetStream.Write($buffer, 0, $count)
           $count = $responseStream.Read($buffer,0,$buffer.length)
           $downloadedBytes = $downloadedBytes + $count
           Write-Progress -id 3 -ParentId 2 -activity "Downloading file '$($url.split('/') | Select -Last 1)'" -status "Downloaded ($([System.Math]::Floor($downloadedBytes/1024))K of $($totalLength)K): " -PercentComplete ((([System.Math]::Floor($downloadedBytes/1024)) / $totalLength)  * 100)
       }

       Write-Progress -id 3 -ParentId 2 -activity "Finished downloading file '$($url.split('/') | Select -Last 1)'"

       $targetStream.Flush()
       $targetStream.Close()
       $targetStream.Dispose()
       $responseStream.Dispose()
       break;
   } catch {
     $strError = $_.Message
     if ($t -ge 9) {
        throw
     }
   }
   Start-Sleep -Milliseconds 500
  }
}

function GetVersionBasedOnThrottle {
    Param(
       [Parameter()]
       [string] $Channel,

       [Parameter()]
       [string] $Version,

       [Parameter()]
       [xml]$currentVerXML
    )

    begin {
        $defaultDisplaySet = 'Throttle','NewestVersion', 'PreviousVersion'

        $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet(‘DefaultDisplayPropertySet’,[string[]]$defaultDisplaySet)
        $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
    }

    Process {
        $results = new-object PSObject[] 0;

        $versionToReturn
        $checkChannel = $Channel
        if($checkChannel -like "FirstReleaseCurrent"){$checkChannel = "InsidersSlow"}

        $historyOfVersionsLink = "http://officecdn.microsoft.com/pr/wsus/releasehistory.cab"

        $destination2 = "$env:TEMP\ReleaseHistory.cab"

        DownloadFile -url $historyOfVersionsLink -targetFile $destination2 | Out-Null
    
        expand $destination2 "$env:TEMP\ReleaseHistory.xml" -f:"ReleaseHistory.xml" | Out-Null
    
        $baseCabFileName2 = "$env:Temp\ReleaseHistory.xml"
        [xml]$vdxml = Get-Content $baseCabFileName2

        $UpdateChannels = $vdxml.ReleaseHistory.UpdateChannel;
        $updates = $UpdateChannels | Where {$_.Name -like $checkChannel}

        #foreach($update in $updates.Update){
        #

            #Write-Host $update.LegacyVersion $update.Latest
                #get base .cab to get current version
                $baseCabFile = $CurrentVersionXML.UpdateFiles.File | ? rename -ne $null
                $url = "$baseURL$($baseCabFile.relativePath)$($baseCabFile.rename)"

                if (!($TargetDirectory)) {
                   $TargetDirectory = $env:TEMP
                }

                $destFolderPath = "$TargetDirectory\$FolderName\Office\Data"
                [system.io.directory]::CreateDirectory($destFolderPath) | Out-Null

                $destination = "$destFolderPath\$($baseCabFile.rename)"

                DownloadFile -url $url -targetFile $destination | Out-Null

                expand $destination $env:TEMP -f:"VersionDescriptor.xml" | Out-Null
                $baseCabFileName = $env:TEMP + "\VersionDescriptor.xml"
                [xml]$vdxml = Get-Content $baseCabFileName
                $throttle = $vdxml.Version.Throttle;

                Remove-Item -Path $baseCabFileName | Out-Null

           $object = New-Object PSObject -Property @{Throttle = $throttle.Value; 
                                                     NewestVersion = $updates.Update[0].LegacyVersion ;
                                                     PreviousVersion = $updates.Update[1].LegacyVersion ;
                                                    }
           $object | Add-Member MemberSet PSStandardMembers $PSStandardMembers
           $results += $object
    
        Remove-Item -Path $baseCabFileName2 | Out-Null
    
        return $results
    }
}

function PurgeOlderVersions([string]$targetDirectory, [int]$numVersionsToKeep, [array]$channels){
    Write-Host "Checking for Older Versions"
    <# write log#>
    $lineNum = Get-CurrentLineNumber    
    $filName = Get-CurrentFileName 
    WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Checking for Older Versions"
                         
    for($k = 0; $k -lt $channels.Count; $k++)
    {
        [array]$totalVersions = @()#declare empty array so each folder can be purged of older versions individually
        [string]$channelName = $channels[$k]
        [string]$shortChannelName = ConvertChannelNameToShortName -ChannelName $channelName
        [string]$branchName = ConvertChannelNameToBranchName -ChannelName $channelName
        [string]$channelName2 = ConvertBranchNameToChannelName -BranchName $channelName

        $folderList = @($channelName, $shortChannelName, $channelName2, $branchName)

        foreach ($folderName in $folderList) {
            $directoryPath = $TargetDirectory.ToString() + '\'+ $folderName +'\Office\Data'

            if (Test-Path -Path $directoryPath) {
               break;
            }
        }

        if (Test-Path -Path $directoryPath) {
            Write-Host "`tChannel: $channelName2"
            <# write log#>
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Channel: $channelName2"
             [bool]$versionsToRemove = $false

            $files = Get-ChildItem $directoryPath  
            Foreach($file in $files)
            {        
                if($file.GetType().Name -eq 'DirectoryInfo')
                {
                    $totalVersions+=$file.Name
                }
            }

            #check if number of versions is greater than number of versions to hold onto, if not, then we don't need to do anything
            if($totalVersions.Length -gt $numVersionsToKeep)
            {
                #sort array in numerical order
                $totalVersions = $totalVersions | Sort-Object 
               
                #delete older versions
                $numToDelete = $totalVersions.Length - $numVersionsToKeep
                for($i = 1; $i -le $numToDelete; $i++)#loop through versions
                {
                     $versionsToRemove = $true
                     $removeVersion = $totalVersions[($i-1)]
                     Write-Host "`t`tRemoving Version: $removeVersion"
                     <# write log#>
                    $lineNum = Get-CurrentLineNumber    
                    $filName = Get-CurrentFileName 
                    WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Removing Version: $removeVersion"
                     
                     Foreach($file in $files)#loop through files
                     {  #array is 0 based

                        if($file.Name.Contains($removeVersion))
                        {                               
                            $folderPath = "$directoryPath\$file"

                             for($t=1;$t -lt 5; $t++) {
                               try {
                                  Remove-Item -Recurse -Force $folderPath
                                  break;
                               } catch {
                                 if ($t -ge 4) {
                                    throw
                                 }
                               }
                             }
                        }
                     }
                }

            }

            if (!($versionsToRemove)) {
                Write-Host "`t`tNo Versions to Remove"
                 <# write log#>
                $lineNum = Get-CurrentLineNumber    
                $filName = Get-CurrentFileName 
                WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "No Versions to Remove"
            }
        }


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

function ConvertChannelNameToBranchName {
    Param(
       [Parameter()]
       [string] $ChannelName
    )
    Process {
       if ($ChannelName.ToLower() -eq "FirstReleaseCurrent".ToLower()) {
         return "FirstReleaseCurrent"
       }
       if ($ChannelName.ToLower() -eq "Current".ToLower()) {
         return "Current"
       }
       if ($ChannelName.ToLower() -eq "FirstReleaseDeferred".ToLower()) {
         return "FirstReleaseBusiness"
       }
       if ($ChannelName.ToLower() -eq "Deferred".ToLower()) {
         return "Business"
       }
       if ($ChannelName.ToLower() -eq "Business".ToLower()) {
         return "Business"
       }
       if ($ChannelName.ToLower() -eq "FirstReleaseBusiness".ToLower()) {
         return "FirstReleaseBusiness"
       }
    }
}

function ConvertBranchNameToChannelName {
    Param(
       [Parameter()]
       [string] $BranchName
    )
    Process {
       if ($BranchName.ToLower() -eq "FirstReleaseCurrent".ToLower()) {
         return "FirstReleaseCurrent"
       }
       if ($BranchName.ToLower() -eq "Current".ToLower()) {
         return "Current"
       }
       if ($BranchName.ToLower() -eq "FirstReleaseDeferred".ToLower()) {
         return "FirstReleaseDeferred"
       }
       if ($BranchName.ToLower() -eq "Deferred".ToLower()) {
         return "Deferred"
       }
       if ($BranchName.ToLower() -eq "Business".ToLower()) {
         return "Deferred"
       }
       if ($BranchName.ToLower() -eq "FirstReleaseBusiness".ToLower()) {
         return "FirstReleaseDeferred"
       }
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

function Check-FileHash {
    Param(
       [Parameter()]
       [string] $FilePath,

       [Parameter()]
       [string] $CabFile
    )
    Process {          
         
        $FileName = Split-Path -Path $FilePath -Leaf -Resolve
        $hashFileName = $FileName.replace("dat","hash")

        Write-Progress -id 3 -ParentId 2 -activity "Checking file hash: $FileName"

        $folderName = [guid]::NewGuid()
        $targetDir = "$env:Temp\$folderName"
        [system.io.directory]::CreateDirectory($targetDir) | Out-Null
                  
        expand $CabFile -f:* $targetDir\ | Out-Null
                                   
        $fileHash = Get-FileHash $FilePath.replace('/','\')
        $providedHash = Get-Content $targetDir\$hashFileName

        if($fileHash.hash -ne $providedHash)
        {
           Write-Progress -id 3 -ParentId 2 -activity "File hashes do not match: $FileName"
           return $false
        }
        else{
           Write-Progress -id 3 -ParentId 2 -activity "File hash check successful: $FileName"
           return $true
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
        $fileName = $_.InvocationInfo.ScriptName.Substring($_.InvocationInfo.ScriptName.LastIndexOf("\")+1)
        WriteToLogFile -LNumber $_.InvocationInfo.ScriptLineNumber -FName $fileName -ActionError $_
    }
}