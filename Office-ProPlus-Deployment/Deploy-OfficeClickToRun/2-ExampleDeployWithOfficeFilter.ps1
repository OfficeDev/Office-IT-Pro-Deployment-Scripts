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
. $scriptPath\EnvironmentalFilter.ps1

$targetFilePath = "$env:temp\configuration.xml"

$SourcePath = $scriptPath
if((Validate-UpdateSource -UpdateSource $SourcePath -ShowMissingFiles $false) -eq $false) {
    $SourcePath = $NULL    
}

#This example will create an Office Deployment Tool (ODT) configuration file and include all of the Languages currently in use on the computer
#from which the script is run.  It will then remove the Version attribute from the XML to ensure the installation gets the latest version
#when updating an existing install and then it will initiate a install

#Generates Configuration Xml based on the local computer
Generate-ODTConfigurationXml -Languages AllInUseLanguages -TargetFilePath $targetFilePath | Out-Null

#Ensure the Version attribute is not set so the install will install the latest version
Set-ODTAdd -TargetFilePath $targetFilePath -Version $NULL -Channel Deferred -SourcePath $SourcePath | Out-Null

#Any workstation in the Paris Office Active Directory Organizational Unit (OU) or sub OU's will have their configuration overrided to set the lanuguage to French
if ((Check-ComputerInOUPath -ContainerPath "OU=Paris" -IncludeSubContainers $true)) {
   $currentProducts = Get-ODTProductToAdd
   foreach ($product in $currentProducts) {
      Set-ODTProductToAdd -TargetFilePath $targetFilePath -ProductId $product.ProductId -LanguageIds @("fr-fr") | Out-Null
   }

   Set-ODTUpdates -TargetFilePath $targetFilePath -Enabled $true -UpdatePath "\\ParisFileServer\OfficeUpdates"
}

#Any workstation in the Tokyo Office Active Directory Organizational Unit (OU) or sub OU's will have their configuration overrided to set the lanuguage to French
if ((Check-ComputerInOUPath -ContainerPath "OU=Tokyo" -IncludeSubContainers $true)) {
   $currentProducts = Get-ODTProductToAdd
   foreach ($product in $currentProducts) {
      Set-ODTProductToAdd -TargetFilePath $targetFilePath -ProductId $product.ProductId -LanguageIds @("ja-jp") | Out-Null
   }

   Set-ODTUpdates -TargetFilePath $targetFilePath -Enabled $true -UpdatePath "\\TokyoFileServer\OfficeUpdates"
}


#Display the Configuration Xml to the Screen
Show-ODTConfiguration $targetFilePath

#Run a Office Deployment Tool (ODT) /configure with the Configuration XML that was generated with this script
Install-OfficeClickToRun -TargetFilePath $targetFilePath

# Configuration.xml file for Click-to-Run for Office 365 products reference. https://technet.microsoft.com/en-us/library/JJ219426.aspx
}
