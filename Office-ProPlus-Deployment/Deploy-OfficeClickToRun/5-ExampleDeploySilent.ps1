#  Office ProPlus Click-To-Run Deployment Script example
#
#  This script demonstrates how utilize the scripts in OfficeDev/Office-IT-Pro-Deployment-Scripts repository together to create
#  Office ProPlus Click-To-Run deployment script that will be adaptive to the configuration of the computer it is run from

Process {
#Importing all required functions
. $PSScriptRoot\Generate-ODTConfigurationXML.ps1
. $PSScriptRoot\Edit-OfficeConfigurationFile.ps1
. $PSScriptRoot\Install-OfficeClickToRun.ps1

$targetFilePath = "configuration.xml"

#This example will create an Office Deployment Tool (ODT) configuration file and include all of the Languages currently in use on the computer
#from which the script is run.  It will then remove the Version attribute from the XML to ensure the installation gets the latest version
#when updating an existing install and then it will initiate a install

#This script additionally sets the "AcceptEULA" to "True" and the display "Level" to "None" so the install is silent.

Generate-ODTConfigurationXml -Languages AllInUseLanguages -TargetFilePath $targetFilePath | Set-ODTAdd -Version $NULL | Set-ODTDisplay -AcceptEULA $true -Level None | Install-OfficeClickToRun

# Configuration.xml file for Click-to-Run for Office 365 products reference. https://technet.microsoft.com/en-us/library/JJ219426.aspx
}