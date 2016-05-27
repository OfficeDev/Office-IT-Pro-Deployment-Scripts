  param(
    [Parameter()]
    [string]$Channel = $null,

    [Parameter()]
    [string]$Bitness = "32",

    [Parameter()]
    [string]$SourceFileFolder = "SourceFiles"
  )

#  Office ProPlus Click-To-Run Deployment Script example
#
#  This script demonstrates how utilize the scripts in OfficeDev/Office-IT-Pro-Deployment-Scripts repository together to create
#  Office ProPlus Click-To-Run deployment script that will be adaptive to the configuration of the computer it is run from

Process {
 $scriptPath = "."

 if ($PSScriptRoot) {
   $scriptPath = $PSScriptRoot
 } else {
   $scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
 }

 $shareFunctionsPath = "$scriptPath\SharedFunctions.ps1"
 if ($scriptPath.StartsWith("\\")) {
 } else {
    if (!(Test-Path -Path $shareFunctionsPath)) {
        throw "Missing Dependency File SharedFunctions.ps1"    
    }
 }
 . $shareFunctionsPath

 $UpdateURLPath = $scriptPath
 if ($SourceFileFolder) {
   if (Test-ItemPathUNC -Path "$UpdateURLPath\$SourceFileFolder") {
      $UpdateURLPath = "$UpdateURLPath\$SourceFileFolder"
   }
 }

 #Importing all required functions
. $scriptPath\Generate-ODTConfigurationXML.ps1
. $scriptPath\Edit-OfficeConfigurationFile.ps1
. $scriptPath\Install-OfficeClickToRun.ps1
. $scriptPath\SharedFunctions.ps1

$UpdateURLPath = Change-UpdatePathToChannel -Channel $Channel -UpdatePath $UpdateURLPath

$targetFilePath = "$env:temp\configuration.xml"

#This example will create an Office Deployment Tool (ODT) configuration file and include all of the Languages currently in use on the computer
#from which the script is run.  It will then remove the Version attribute from the XML to ensure the installation gets the latest version
#when updating an existing install and then it will initiate a install

Generate-ODTConfigurationXml -Languages AllInUseLanguages -TargetFilePath $targetFilePath | Set-ODTAdd -Version $NULL | Set-ODTDisplay -Level None -AcceptEULA $true 

$languages = Get-XMLLanguages -Path $targetFilePath

if (Test-UpdateSource -UpdateSource $UpdateURLPath -OfficeLanguages $languages) {
   Set-ODTAdd -TargetFilePath $targetFilePath -SourcePath $UpdateURLPath
}

if (($Bitness -eq "32") -or ($Bitness -eq "x86")) {
    Set-ODTAdd -TargetFilePath $targetFilePath -Bitness 32
} else {
    Set-ODTAdd -TargetFilePath $targetFilePath -Bitness 64
}

Install-OfficeClickToRun -TargetFilePath $targetFilePath

# Configuration.xml file for Click-to-Run for Office 365 products reference. https://technet.microsoft.com/en-us/library/JJ219426.aspx
}