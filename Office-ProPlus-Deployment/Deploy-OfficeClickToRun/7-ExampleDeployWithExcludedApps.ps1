#  Office ProPlus Click-To-Run Deployment Script example
#
#  This script demonstrates how utilize the scripts in OfficeDev/Office-IT-Pro-Deployment-Scripts repository together to create
#  Office ProPlus Click-To-Run deployment script that will be adaptive to the configuration of the computer it is run from

Process {
 $scriptPath = "."

 if ($PSScriptRoot) {
   $scriptPath = $PSScriptRoot
 } else {
   $scriptPath = (Get-Item -Path ".\").FullName
 }

#Importing all required functions
. $scriptPath\Generate-ODTConfigurationXML.ps1
. $scriptPath\Edit-OfficeConfigurationFile.ps1
. $scriptPath\Install-OfficeClickToRun.ps1

$targetFilePath = "$env:temp\configuration.xml"

$SourcePath = $scriptPath
if((Validate-UpdateSource -UpdateSource $SourcePath -ShowMissingFiles $false) -eq $false) {
    $SourcePath = $NULL    
}

#This example will create an Office Deployment Tool (ODT) configuration file and include all of the Languages currently in use on the computer
#from which the script is run.  It will then remove the Version attribute from the XML to ensure the installation gets the latest version and 
#will set the Channel to 'Deferred'.  It will then detect if O365ProPlusRetail or O365BusinessRetail is in the configuration file and if so
#it will add Lync and Groove to the excluded apps. It will then initiate the Office installation.

Generate-ODTConfigurationXml -Languages AllInUseLanguages -TargetFilePath $targetFilePath | Set-ODTAdd -Version $NULL -Channel Deferred -SourcePath $SourcePath | Out-Null

if ((Get-ODTProductToAdd -TargetFilePath $targetFilePath -ProductId O365ProPlusRetail)) {
     Set-ODTProductToAdd -ProductId "O365ProPlusRetail" -TargetFilePath $targetFilePath -ExcludeApps ("Lync", "Groove") | Out-Null
}

if ((Get-ODTProductToAdd -TargetFilePath $targetFilePath -ProductId O365BusinessRetail)) {
     Set-ODTProductToAdd -ProductId "O365BusinessRetail" -TargetFilePath $targetFilePath -ExcludeApps ("Lync", "Groove") | Out-Null
}

Install-OfficeClickToRun -TargetFilePath $targetFilePath


# Configuration.xml file for Click-to-Run for Office 365 products reference. https://technet.microsoft.com/en-us/library/JJ219426.aspx
}

