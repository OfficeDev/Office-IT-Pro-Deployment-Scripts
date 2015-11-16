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

function Download-OfficeBranch{
<#
.SYNOPSIS
Downloads cab files for specified branches, versions, bitness, and languages
.DESCRIPTION
Downloads cab files for specified branches, versions, bitness, and languages
.PARAMETER version
The version number you wish to download. For example: 16.0.6228.1010
.PARAMETER baseDestination
Required. Where all the branches will be downloaded. Each branch then goes into a folder of the same name as the branch.
.PARAMETER languages
Array of Microsoft language codes. Will throw error if provided values don't match the validation set. Defaults to "en-us"
("en-us","ar-sa","bg-bg","zh-cn","zh-tw","hr-hr","cs-cz","da-dk","nl-nl","et-ee","fi-fi","fr-fr","de-de","el-gr","he-il","hi-in","hu-hu","id-id","it-it",
"ja-jp","kk-kh","ko-kr","lv-lv","lt-lt","ms-my","nb-no","pl-pl","pt-br","pt-pt","ro-ro","ru-ru","sr-latn-rs","sk-sk","sl-si","es-es","sv-se","th-th",
"tr-tr","uk-ua")
.PARAMETER bitness
v32, v64, or Both. What bitness of office you wish to download. Defaults to Both.
.PARAMETER branches
An array of the branches you wish to download. Defaults to all available branches (CMValidation currently not available)
.Example
Download-OfficeBranch -baseDestination "C:\Users\Public\Documents\"
Default downloads all available branches of the most recent version for both bitnesses into public documents. Downloads the english language pack.
.Example
Download-OfficeBranch -baseDestination "C:\Users\Public\Documents\" -SourceVersion "14.0" -TargetVersion "15.0"
Copy the office 14.0 (2010) policies within 'Office Settings' to office 15.0 (2013) policies within 'Office Settings'
.Link
https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts
#>

Param(
    [Parameter()]
    [string] $version,

    [Parameter(Mandatory=$true)]
    [string] $baseDestination,

    [Parameter()]
    [ValidateSet("en-us","ar-sa","bg-bg","zh-cn","zh-tw","hr-hr","cs-cz","da-dk","nl-nl","et-ee","fi-fi","fr-fr","de-de","el-gr","he-il","hi-in","hu-hu","id-id","it-it",
                "ja-jp","kk-kh","ko-kr","lv-lv","lt-lt","ms-my","nb-no","pl-pl","pt-br","pt-pt","ro-ro","ru-ru","sr-latn-rs","sk-sk","sl-si","es-es","sv-se","th-th",
                "tr-tr","uk-ua")]
    [string[]] $languages = ("en-us"),

    [Parameter()]
    [Bitness] $bitness = 0,

    [Parameter()]
    [OfficeBranch[]] $branches = (0, 1, 2, 3)#, 4)
)
$numberOfFiles = (($branches.Count + 1) * ((($languages.Count + 1)*3) + 5))

$webclient = New-Object System.Net.WebClient
$XMLFilePath = "$env:TEMP/ofl.cab"
$XMLDownloadURL = "http://officecdn.microsoft.com/pr/wsus/ofl.cab"
$webclient.DownloadFile($XMLDownloadURL,$XMLFilePath)
if($bitness -eq [Bitness]::Both -or $bitness -eq [Bitness]::v32){
    $32XMLFileName = "o365client_32bit.xml"
    expand $XMLFilePath $env:TEMP -f:$32XMLFileName | Out-Null
    $32XMLFilePath = $env:TEMP + "\o365client_32bit.xml"
    [xml]$32XML = Get-Content $32XMLFilePath
    $xmlArray = ($32XML)
}

if($bitness -eq [Bitness]::Both -or $bitness -eq [Bitness]::v64){
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
Write-Progress -Activity "Downloading Branch Files" -status "Beginning" -percentComplete ($j / $numberOfFiles *100)

#loop to download files
$xmlArray | %{
    $CurrentVersionXML = $_
    
    #loop for each branch
    $branches | %{
        $currentBranch = $_
        $baseURL = $CurrentVersionXML.UpdateFiles.baseURL | ? branch -eq $_.ToString() | %{$_.URL};
        if(!(Test-Path "$baseDestination\$($_.ToString())\")){
            New-Item -Path "$baseDestination\$($_.ToString())\"  -ItemType directory -Force
        }

        if([String]::IsNullOrWhiteSpace($version)){
            #get base .cab to get current version
            $webclient = New-Object System.Net.WebClient
            $baseCabFile = $CurrentVersionXML.UpdateFiles.File | ? rename -ne $null
            $url = "$baseURL$($baseCabFile.relativePath)$($baseCabFile.rename)"
            $destination = "$baseDestination\$($_.ToString())\$($baseCabFile.rename)"
            $webclient.DownloadFile($url,$destination)

            expand $destination $env:TEMP -f:"VersionDescriptor.xml" | Out-Null
            $baseCabFileName = $env:TEMP + "\VersionDescriptor.xml"
            [xml]$vdxml = Get-Content $baseCabFileName
            $currentVersion = $vdxml.Version.Available.Build;
            Remove-Item -Path $baseCabFileName
        }else{
            $currentVersion = $version
        }

        #basic files
        $CurrentVersionXML.UpdateFiles.File | ? language -eq "0" | 
        %{
            $webclient = New-Object System.Net.WebClient
            $name = $_.name -replace "`%version`%", $currentVersion
            $relativePath = $_.relativePath -replace "`%version`%", $currentVersion
            $url = "$baseURL$relativePath$name"
            $destination = "$baseDestination\$($currentBranch.ToString())\$name"
            $webclient.DownloadFile($url,$destination)
            $j = $j + 1
            Write-Progress -Activity "Downloading Branch Files" -status "Branch: $($currentBranch.ToString())" -percentComplete ($j / $numberOfFiles *100)
        }

        #language files
        $languages | 
        %{
            #LANGUAGE LOGIC HERE
            $languageId  = [globalization.cultureinfo]::GetCultures("allCultures") | ? Name -eq $_ | %{$_.LCID}
            $CurrentVersionXML.UpdateFiles.File | ? language -eq $languageId | 
            %{
                $webclient = New-Object System.Net.WebClient
                $name = $_.name -replace "`%version`%", $currentVersion
                $relativePath = $_.relativePath -replace "`%version`%", $currentVersion
                $url = "$baseURL$relativePath$name"
                $destination = "$baseDestination\$($currentBranch.ToString())\$name"
                $webclient.DownloadFile($url,$destination)
                $j = $j + 1
                Write-Progress -Activity "Downloading Branch Files" -status "Branch: $($currentBranch.ToString())" -percentComplete ($j / $numberOfFiles *100)
            }
        }
    }

}
}