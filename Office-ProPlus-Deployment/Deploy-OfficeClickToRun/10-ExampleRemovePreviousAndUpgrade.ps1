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

#Sets whether to use Volume Licensing for Project and Visio
$UseVolumeLicensing = $false

#Importing all required functions
. $scriptPath\Generate-ODTConfigurationXML.ps1
. $scriptPath\Install-OfficeClickToRun.ps1
. $scriptPath\Remove-PreviousOfficeInstalls.ps1
. $scriptPath\Remove-OfficeClickToRun.ps1
. $scriptPath\SharedFunctions.ps1
. $scriptPath\Edit-OfficeConfigurationFile.ps1

$targetFilePath = "$env:temp\configuration.xml"

#This example will create an Office Deployment Tool (ODT) configuration file and include all of the Languages currently in use on the computer
#from which the script is run.  It will then remove the Version attribute from the XML to ensure the installation gets the latest version
#when updating an existing install and then it will initiate a install

#This script additionally sets the "AcceptEULA" to "True" and the display "Level" to "None" so the install is silent.

$officeProducts = Get-OfficeVersion -ShowAllInstalledProducts | Select *

$Office2016C2RExists = $officeProducts | Where {$_.ClickToRun -eq $true -and $_.Version -like '16.*' }

$SourcePath = $scriptPath
if((Validate-UpdateSource -UpdateSource $SourcePath -ShowMissingFiles $false) -eq $false) {
    $SourcePath = $NULL    
}

if ($Office2016C2RExists) {
  Write-Host "Office 2016 Click-To-Run is already installed"
} else {
    if (!(Test-Path -Path $targetFilePath)) {
       Generate-ODTConfigurationXml -Languages AllInUseLanguages -TargetFilePath $targetFilePath | Set-ODTAdd -Version $NULL -SourcePath $SourcePath -Channel Deferred | Set-ODTDisplay -Level None -AcceptEULA $true | Out-Null

       $products = Get-ODTProductToAdd -TargetFilePath $targetFilePath -All
       if ($products) { $languages = $products.Languages } else { $languages = @("en-us") }
       $visioAdded = $products | Where { $_.ProductID -like 'VisioProRetail' }
       $projectAdded = $products | Where { $_.ProductID -like 'ProjectProRetail' }
       
       $VisioPro = $officeProducts | Where { $_.DisplayName -like '*Visio Professional*' -and $_.ClickToRun -eq $false }
       $VisioStd = $officeProducts | Where { $_.DisplayName -like '*Visio Standard*' -and $_.ClickToRun -eq $false }
       $ProjectPro = $officeProducts | Where { $_.DisplayName -like '*Project Professional*' -and $_.ClickToRun -eq $false }
       $ProjectStd = $officeProducts | Where { $_.DisplayName -like '*Project Standard*' -and $_.ClickToRun -eq $false }

       if ($UseVolumeLicensing) {
           if ($visioAdded) { Remove-ODTProductToAdd -ProductId 'VisioProRetail' -TargetFilePath $targetFilePath }
           if ($projectAdded) { Remove-ODTProductToAdd -ProductId 'ProjectProRetail' -TargetFilePath $targetFilePath }

           if ($VisioPro.Count -gt 0) { Add-ODTProductToAdd -ProductId VisioProXVolume -TargetFilePath $targetFilePath -LanguageIds $languages | Out-Null }
           if ($VisioStd.Count -gt 0) { Add-ODTProductToAdd -ProductId VisioStdXVolume -TargetFilePath $targetFilePath -LanguageIds $languages | Out-Null }
           if ($ProjectPro.Count -gt 0) { Add-ODTProductToAdd -ProductId ProjectProXVolume -TargetFilePath $targetFilePath -LanguageIds $languages | Out-Null }
           if ($ProjectStd.Count -gt 0) { Add-ODTProductToAdd -ProductId ProjectStdXVolume -TargetFilePath $targetFilePath -LanguageIds $languages | Out-Null }
       }
    }else {
        Set-ODTAdd -SourcePath $SourcePath -TargetFilePath $TargetFilePath | Out-Null
    }

    Remove-OfficeClickToRun 

    Remove-PreviousOfficeInstalls

    Install-OfficeClickToRun -TargetFilePath $targetFilePath
}

# Configuration.xml file for Click-to-Run for Office 365 products reference. https://technet.microsoft.com/en-us/library/JJ219426.aspx
}
