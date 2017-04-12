param(
 [Parameter()]
 [string]$OfficeDeploymentPath
)

Begin {
 Set-Location $OfficeDeploymentPath
}

Process {
 $scriptPath = "."

 # Importing required functions
 . $scriptPath\Generate-ODTConfigurationXML.ps1
 . $scriptPath\Edit-OfficeConfigurationFile.ps1
 . $scriptPath\SharedFunctions.ps1
 . $scriptPath\Remove-PreviousOfficeInstalls.ps1
  
 #--------------------------------------------------------------------------------------
 #   Customize the parameters - Modify the variables below to customize this script
 #--------------------------------------------------------------------------------------

 # Available list of ProductsToRemove:  AllOfficeProducts,MainOfficeProduct,Visio,Project
 [string[]]$ProductsToRemove = "AllOfficeProducts"
 
 [bool]$RemoveClickToRunVersions = $false
 
 [bool]$Remove2016Installs = $false
 
 [bool]$Force = $false
 
 [bool]$KeepUserSettings = $true
 
 [bool]$KeepLync = $false

 [bool]$NoReboot = $false

 [string]$LogFilePath = "$env:TEMP\RemovePreviousOfficeInstall.log"

 #-------------------------------------------------------------------------------------

 # Remove the products
 Remove-PreviousOfficeInstalls -ProductsToRemove $ProductsToRemove -RemoveClickToRunVersions $RemoveClickToRunVersions -Remove2016Installs $Remove2016Installs -Force $Force -KeepUserSettings $KeepUserSettings -KeepLync $KeepLync -NoReboot $NoReboot -LogFilePath $LogFilePath

}