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
. $scriptPath\SharedFunctions.ps1
. $scriptPath\Edit-OfficeConfigurationFile.ps1
. $scriptPath\Install-OfficeClickToRun.ps1
. $scriptPath\Remove-OfficeClickToRun.ps1

$targetFilePath = "$env:temp\configuration.xml"

$SourcePath = $scriptPath
if((Validate-UpdateSource -UpdateSource $SourcePath -ShowMissingFiles $false) -eq $false) {
    $SourcePath = $NULL    
}

#This example will create an Office Deployment Tool (ODT) configuration file and include all of the Languages currently in use on the computer
#from which the script is run. It will then remove the existing Office 2016 Click-To-Run installation and then it will then remove the Version attribute from the XML to ensure the installation gets the latest version
#when updating an existing install and then it will initiate a install of Office 2013 Click-To-Run

Generate-ODTConfigurationXml -Languages AllInUseLanguages -TargetFilePath $targetFilePath | Remove-OfficeClickToRun | Set-ODTAdd -Version $NULL -SourcePath $SourcePath | Install-OfficeClickToRun -OfficeVersion Office2013

# Configuration.xml file for Click-to-Run for Office 365 products reference. https://technet.microsoft.com/en-us/library/JJ219426.aspx

}
