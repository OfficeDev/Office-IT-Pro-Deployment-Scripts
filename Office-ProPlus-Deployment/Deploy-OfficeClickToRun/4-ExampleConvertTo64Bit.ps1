#  Office ProPlus Click-To-Run Deployment Script example
#
#  This script demonstrates how utilize the scripts in OfficeDev/Office-IT-Pro-Deployment-Scripts repository together to create
#  Office ProPlus Click-To-Run deployment script that will Convert existing 32-Bit Click-To-Run installs to 64-bit with the same 
#  Products and Languages that are currently installed

Process {
 $scriptPath = "."

 if ($PSScriptRoot) {
   $scriptPath = $PSScriptRoot
 } else {
   $scriptPath = (Get-Item -Path ".\").FullName
 }

#Importing all required functions - These files must be in the same directory as this script
. $scriptPath\Generate-ODTConfigurationXML.ps1
. $scriptPath\SharedFunctions.ps1
. $scriptPath\Edit-OfficeConfigurationFile.ps1
. $scriptPath\Install-OfficeClickToRun.ps1
. $scriptPath\Remove-OfficeClickToRun.ps1
. $scriptPath\Get-OfficeVersion.ps1
. $scriptPath\Remove-PreviousOfficeInstalls.ps1

$targetFilePath = "$env:temp\configuration.xml"

$SourcePath = $scriptPath
if((Validate-UpdateSource -UpdateSource $SourcePath -ShowMissingFiles $false) -eq $false) {
    $SourcePath = $NULL    
}

#This example will detect the current install of Office that is currently installed. If the current install of Office Click-To-Run 32-bit it will
#then generate a Configuration XML based on the current configuration It will then remove the Version attribute from the XML to ensure the installation gets the latest version
#and change the configuration XML to 64-Bit.  It will remove the existing install of Office Click-To-Run and resinstall Office Click-To-Run with the 64-Bit version

$installOffice = $true

$officeC2R = getCTRConfig
if ($officeC2R) {
    if ($officeC2R.Platform -eq "32") {
        if (!(Test-Path -Path $targetFilePath)) {
            Generate-ODTConfigurationXml -Languages CurrentOfficeLanguages -TargetFilePath $targetFilePath | Out-Null
        }
    } else {
      $installOffice = $false
    }
}

if ($installOffice) {
  if (Test-Path -Path $targetFilePath) {
      $PinnedStartMenuApps = GetPinnedStartMenuApps
      if($PinnedStartMenuApps -eq $NULL){
          $PinnedStartMenuApps = "None"
      }

      Remove-OfficeClickToRun -TargetFilePath $targetFilePath

      Remove-PreviousOfficeInstalls

      Set-ODTAdd -TargetFilePath $targetFilePath -Version $NULL -Bitness 64 -SourcePath $SourcePath | Out-Null
      Set-ODTDisplay -TargetFilePath $targetFilePath -Level None -AcceptEULA $true | Out-Null
      Set-ODTLogging -TargetFilePath $targetFilePath -Path $env:temp -Level Standard | Out-Null 

      $updates = Get-ODTUpdates -TargetFilePath $targetFilePath
      $addNode = Get-ODTAdd -TargetFilePath $targetFilePath

      $UpdatePath = $NULL
      if ($updates) {
         $UpdatePath = $updates.UpdatePath
         if ($UpdatePath -like '*officecdn.microsoft.com*') {
             $UpdatePath = $NULL
         }
      }

      Set-ODTUpdates -TargetFilePath $targetFilePath -Channel $addNode.Channel -Enabled $true -UpdatePath $UpdatePath | Out-Null

      Restart-ExplorerExe

      Install-OfficeClickToRun -TargetFilePath $targetFilePath -PinToStartMenu $PinnedStartMenuApps
      
  }
}

# Configuration.xml file for Click-to-Run for Office 365 products reference. https://technet.microsoft.com/en-us/library/JJ219426.aspx
}
