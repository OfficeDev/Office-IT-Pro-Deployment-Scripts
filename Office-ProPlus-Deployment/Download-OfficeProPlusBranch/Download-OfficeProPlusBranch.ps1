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

Add-Type -TypeDefinition $enumDef

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

Add-Type -TypeDefinition $enumDef

function Download-OfficeProPlusBranch{
<#
.SYNOPSIS
Downloads each Office ProPlus Branch with installation files
.DESCRIPTION
This script will dynamically downloaded the most current Office ProPlus version for each deployment Branch
.PARAMETER Version
The version number you wish to download. For example: 16.0.6228.1010
.PARAMETER TargetDirectory
Required. Where all the branches will be downloaded. Each branch then goes into a folder of the same name as the branch.
.PARAMETER Languages
Array of Microsoft language codes. Will throw error if provided values don't match the validation set. Defaults to "en-us"
("en-us","ar-sa","bg-bg","zh-cn","zh-tw","hr-hr","cs-cz","da-dk","nl-nl","et-ee","fi-fi","fr-fr","de-de","el-gr","he-il","hi-in","hu-hu","id-id","it-it",
"ja-jp","kk-kh","ko-kr","lv-lv","lt-lt","ms-my","nb-no","pl-pl","pt-br","pt-pt","ro-ro","ru-ru","sr-latn-rs","sk-sk","sl-si","es-es","sv-se","th-th",
"tr-tr","uk-ua")
.PARAMETER Bitness
v32, v64, or Both. What bitness of office you wish to download. Defaults to Both.
.PARAMETER OverWrite
If this parameter is specified then existing files will be overwritten.
.PARAMETER Branches
An array of the branches you wish to download. Defaults to all available branches (CMValidation currently not available)
.Example
Download-OfficeBranch -baseDestination "\\server\updateshare"
Default downloads all available branches of the most recent version for both bitnesses into an update source. Downloads the English language pack by default if language is not specified.
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
                "ja-jp","kk-kh","ko-kr","lv-lv","lt-lt","ms-my","nb-no","pl-pl","pt-br","pt-pt","ro-ro","ru-ru","sr-latn-rs","sk-sk","sl-si","es-es","sv-se","th-th",
                "tr-tr","uk-ua")]
    [string[]] $Languages = ("en-us"),

    [Parameter()]
    [Bitness] $Bitness = 0,

    [Parameter()]
    [bool] $OverWrite = $false,

    [Parameter()]
    [OfficeBranch[]] $Branches = (0, 1, 2, 3)#, 4)
)

$numberOfFiles = (($Branches.Count) * ((($Languages.Count + 1)*3) + 5))

$webclient = New-Object System.Net.WebClient
$XMLFilePath = "$env:TEMP/ofl.cab"
$XMLDownloadURL = "http://officecdn.microsoft.com/pr/wsus/ofl.cab"
$webclient.DownloadFile($XMLDownloadURL,$XMLFilePath)

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
$BranchCount = $Branches.Count * 2

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
    $Branches | %{
        $currentBranch = $_
        $b++

        Write-Progress -id 1 -Activity "Downloading Branch" -status "Branch: $($currentBranch.ToString()) : $currentBitness" -percentComplete ($b / $BranchCount *100) 

        Write-Host "`tDownloading Branch: $currentBranch"

        $baseURL = $CurrentVersionXML.UpdateFiles.baseURL | ? branch -eq $_.ToString() | %{$_.URL};
        if(!(Test-Path "$TargetDirectory\$($_.ToString())\")){
            New-Item -Path "$TargetDirectory\$($_.ToString())\" -ItemType directory -Force | Out-Null
        }
        if(!(Test-Path "$TargetDirectory\$($_.ToString())\Office")){
            New-Item -Path "$TargetDirectory\$($_.ToString())\Office" -ItemType directory -Force | Out-Null
        }
        if(!(Test-Path "$TargetDirectory\$($_.ToString())\Office\Data")){
            New-Item -Path "$TargetDirectory\$($_.ToString())\Office\Data" -ItemType directory -Force | Out-Null
        }

        if([String]::IsNullOrWhiteSpace($Version)){
            #get base .cab to get current version
            $webclient = New-Object System.Net.WebClient
            $baseCabFile = $CurrentVersionXML.UpdateFiles.File | ? rename -ne $null
            $url = "$baseURL$($baseCabFile.relativePath)$($baseCabFile.rename)"
            $destination = "$TargetDirectory\$($_.ToString())\Office\Data\$($baseCabFile.rename)"

            $webclient.DownloadFile($url,$destination)

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

        if(!(Test-Path "$TargetDirectory\$($_.ToString())\Office\Data\$currentVersion")){
            New-Item -Path "$TargetDirectory\$($_.ToString())\Office\Data\$currentVersion" -ItemType directory -Force | Out-Null
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
            $destination = "$TargetDirectory\$($currentBranch.ToString())$relativePath$name"

            if (!(Test-Path -Path $destination) -or $OverWrite) {
               DownloadFile -url $url -targetFile $destination
            }

            $j = $j + 1
            Write-Progress -id 2 -ParentId 1 -Activity "Downloading Branch Files" -status "Branch: $($currentBranch.ToString())" -percentComplete ($j / $numberOfFiles *100)
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
                $destination = "$TargetDirectory\$($currentBranch.ToString())$relativePath$name"

                if (!(Test-Path -Path $destination) -or $OverWrite) {
                   DownloadFile -url $url -targetFile $destination
                }

                $j = $j + 1
                Write-Progress -id 2 -ParentId 1 -Activity "Downloading Branch Files" -status "Branch: $($currentBranch.ToString())" -percentComplete ($j / $numberOfFiles *100)
            }
        }


    }

}
}

function DownloadFile($url, $targetFile) {

   $uri = New-Object "System.Uri" "$url"

   $request = [System.Net.HttpWebRequest]::Create($uri)

   $request.set_Timeout(15000) #15 second timeout

   $response = $request.GetResponse()

   $totalLength = [System.Math]::Floor($response.get_ContentLength()/1024)

   $responseStream = $response.GetResponseStream()

   $targetStream = New-Object -TypeName System.IO.FileStream -ArgumentList $targetFile, Create

   $buffer = new-object byte[] 10KB

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

}