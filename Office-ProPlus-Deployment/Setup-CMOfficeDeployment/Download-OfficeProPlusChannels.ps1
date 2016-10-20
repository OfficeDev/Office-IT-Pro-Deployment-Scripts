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
                "tr-tr","uk-ua")]
    [string[]] $Languages = ("en-us"),

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
    [bool] $IncludeChannelInfo = $false
)

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
      
$numberOfFiles = (($BranchesOrChannels.Count) * ((($Languages.Count + 1)*3) + 5))

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

            #loop for each branch
            $BranchesOrChannels | %{
                $currentBranch = $_
                $b++

                Write-Progress -id 1 -Activity "Downloading Channel" -status "Channel: $($currentBranch.ToString()) : $currentBitness" -percentComplete ($b / $BranchCount *100) 
                Write-Host "`tDownloading Channel: $currentBranch"

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
                      return 
                    }
                }

                if(!(Test-Path "$TargetDirectory\$FolderName\Office\Data\$currentVersion")){
                    New-Item -Path "$TargetDirectory\$FolderName\Office\Data\$currentVersion" -ItemType directory -Force | Out-Null
                }

                $numberOfFiles = 0
                $j = 0

                $CurrentVersionXML.UpdateFiles.File | ? language -eq "0" | 
                %{
                   $numberOfFiles ++
                }

                $Languages | 
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
                    $destination = "$TargetDirectory\$FolderName$relativePath$name"

                    if (!(Test-Path -Path $destination) -or $OverWrite) {
                       DownloadFile -url $url -targetFile $destination
                    }

                    $j = $j + 1
                    Write-Progress -id 2 -ParentId 1 -Activity "Downloading Channel Files" -status "Channel: $($currentBranch.ToString())" -percentComplete ($j / $numberOfFiles *100)
                }

                #language files
                $Languages | 
                %{
                    #LANGUAGE LOGIC HERE
                    $languageId  = [globalization.cultureinfo]::GetCultures("allCultures") | ? Name -eq $_ | %{$_.LCID}
                    $CurrentVersionXML.UpdateFiles.File | ? language -eq $languageId | 
                    %{
                        $name = $_.name -replace "`%version`%", $currentVersion
                        $relativePath = $_.relativePath -replace "`%version`%", $currentVersion
                        $url = "$baseURL$relativePath$name"
                        $destination = "$TargetDirectory\$FolderName$relativePath$name"

                        if (!(Test-Path -Path $destination) -or $OverWrite) {
                           DownloadFile -url $url -targetFile $destination
                        }

                        $j = $j + 1
                        Write-Progress -id 2 -ParentId 1 -Activity "Downloading Channel Files" -status "Channel: $($currentBranch.ToString())" -percentComplete ($j / $numberOfFiles *100)
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
    }

    if($downloadSuccess){#if download succeeds, breaks out of loop
        break
    }

}#end of for loop

PurgeOlderVersions $TargetDirectory $NumVersionsToKeep $BranchesOrChannels

}

function DownloadFile($url, $targetFile) {

  for($t=1;$t -lt 10; $t++) {
   try {
       $uri = New-Object "System.Uri" "$url"
       $request = [System.Net.HttpWebRequest]::Create($uri)
       $request.set_Timeout(3000) #15 second timeout

       $response = $request.GetResponse()
       $totalLength = [System.Math]::Floor($response.get_ContentLength()/1024)
       $responseStream = $response.GetResponseStream()
       $targetStream = New-Object -TypeName System.IO.FileStream -ArgumentList $targetFile, Create
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

function PurgeOlderVersions([string]$targetDirectory, [int]$numVersionsToKeep, [array]$channels){
    Write-Host "Checking for Older Versions"
                         
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
