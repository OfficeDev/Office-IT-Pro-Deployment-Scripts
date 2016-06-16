#  Office ProPlus Click-To-Run Deployment Script example
#
#  This script demonstrates how utilize the scripts in OfficeDev/Office-IT-Pro-Deployment-Scripts repository together to deploy the
#  OneDrive for Business Next Generation Sync Client.

Process {
 $scriptPath = "."

 if ($PSScriptRoot) {
   $scriptPath = $PSScriptRoot
 } else {
   $scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
 }

#Importing all required functions
. $scriptPath\Install-OneDriveForBusiness.ps1

#This example will add the DefaultToBusinessFRE and EnableAddAccounts registry keys, and OneDrive.exe will install silently in the 
#background. 

Install-OneDriveForBusiness -DeploymentType All -Visibility Silent

# Configuration.xml file for Click-to-Run for Office 365 products reference. https://technet.microsoft.com/en-us/library/JJ219426.aspx
}