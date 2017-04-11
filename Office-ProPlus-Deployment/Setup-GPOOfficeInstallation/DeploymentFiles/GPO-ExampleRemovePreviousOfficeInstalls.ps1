param(
[Parameter()]
[string]$OfficeDeploymentPath
)

# Deploy Office 365 ProPlus using Group Policy
Begin {
    Set-Location $OfficeDeploymentPath
}

Process {
 $scriptPath = "."

 # Importing required functions
 . $scriptPath\Generate-ODTConfigurationXML.ps1
 . $scriptPath\Edit-OfficeConfigurationFile.ps1
 . $scriptPath\Remove-PreviousOfficeInstalls.ps1
 . $scriptPath\SharedFunctions.ps1

 #--------------------------------------------------------------------------------------
 #   Customize the parameters - Modify the variables below to customize this script
 #--------------------------------------------------------------------------------------

 [bool]$RemoveClickToRunVersions = $false
 
 [bool]$Remove2016Installs = $false
 
 [bool]$Force = $true
 
 [bool]$KeepUserSettings = $true
 
 [bool]$KeepLync = $false

 [bool]$NoReboot = $false

 # Available list of products:  AllOfficeProducts,MainOfficeProduct,Visio,Project
 [string[]]$ProductsToRemove = "AllOfficeProducts" 

 #-------------------------------------------------------------------------------------

 # Remove the products
 Remove-PreviousOfficeInstalls -RemoveClickToRunVersions $RemoveClickToRunVersions -Remove2016Installs $Remove2016Installs -Force $Force -KeepUserSettings $KeepUserSettings -KeepLync $KeepLync -NoReboot $NoReboot -ProductsToRemove $ProductsToRemove

}