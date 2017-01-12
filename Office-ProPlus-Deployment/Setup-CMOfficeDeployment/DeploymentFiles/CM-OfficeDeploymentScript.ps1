  param(
    [Parameter()]
    [string]$Channel = $null,

    [Parameter()]
    [string]$Bitness = "32",

    [Parameter()]
    [string]$SourceFileFolder = "SourceFiles",

    [Parameter()]
    [string]$Languages,

    [Parameter()]
    [string]$ExcludedApps,

    [Parameter()]
    [string]$AdditionalApps
  )

#  Deploy Office 365 ProPlus using ConfigMgr

Process {
 $targetFilePath = "$env:temp\configuration.xml"

 [string]$scriptPath = "."

 if ($PSScriptRoot) {
    $scriptPath = $PSScriptRoot
 } else {
   $scriptPath = (Get-Item -Path ".\").FullName
 }

 if (Get-OfficeC2RVersion) { Write-Host "Office 365 ProPlus Already Installed" }

 #Importing required functions
 . $scriptPath\Generate-ODTConfigurationXML.ps1
 . $scriptPath\Edit-OfficeConfigurationFile.ps1
 . $scriptPath\Install-OfficeClickToRun.ps1
 . $scriptPath\SharedFunctions.ps1

 $SourcePath = $scriptPath
 if((Validate-UpdateSource -UpdateSource $SourcePath -ShowMissingFiles $false) -eq $false) {
    $SourcePath = $NULL    
 } else {
    $ChannelShortName = ConvertChannelNameToShortName -ChannelName $Channel
    $SourcePath = $SourcePath + "\" + $SourceFileFolder + "\" + $ChannelShortName
 }

 $UpdateURLPath = Locate-UpdateSource -Channel $Channel -UpdateURLPath $scriptPath -SourceFileFolder $SourceFileFolder
 Generate-ODTConfigurationXml -Languages AllInUseLanguages -TargetFilePath $targetFilePath | Set-ODTAdd -Sourcepath $SourcePath -Version $NULL -Channel $Channel | Set-ODTUpdates -Channel $Channel -UpdatePath $UpdateURLPath | Set-ODTDisplay -Level None -AcceptEULA $true  | Out-Null
 Update-ConfigurationXml -TargetFilePath $targetFilePath -UpdateURLPath $UpdateURLPath -Channel $Channel
 
 if(!$Languages){
    $languages = Get-XMLLanguages -Path $TargetFilePath
 } else {
    if($Languages -match ","){
        $Languages = $Languages.Split(",")
    }
 }

 #------------------------------------------------------------------------------------------------------------
 #   Customize Deployment Script - Uncomment and modify the code below to customize this deployment script
 #------------------------------------------------------------------------------------------------------------

  #### ------- Exclude Applications ------- ####
  # Exclude-Applications -TargetFilePath $targetFilePath -ExcludeApps @("Access","Excel","Groove","InfoPath","Lync","OneDrive","OneNote","Outlook","PowerPoint","Project","Publisher","SharePointDesigner","Visio","Word")
 

  #### ------- Add an additional Product Sku ------- ####
  # Add-ProductSku -TargetFilePath $targetFilePath -Languages $languages -ProductIDs O365ProPlusRetail,O365BusinessRetail,VisioProRetail,ProjectProRetail


  #### ------- Remove an additional Product Sku ------- ####
  # Remove-ProductSku -TargetFilePath $targetFilePath -Languages $languages -ProductIDs O365ProPlusRetail,O365BusinessRetail,VisioProRetail,ProjectProRetail


  #### ------- Add languages to all Product Skus in the Configuration Xml File ------- ####
  # Add-ProductLanguage -TargetFilePath $targetFilePath -ProductIDs All -Languages fr-fr,it-it 


  #### ------- Remove languages from all Product Skus in the Configuration Xml File ------- ####
  # Remove-ProductLanguage -TargetFilePath $targetFilePath -ProductIDs All -Languages fr-fr,it-it 


  #### ------- Set the display to Full so the installation   ------- ####
  # Set-ODTDisplay -TargetFilePath $targetFilePath -Level Full -AcceptEULA $true


  #### ------- Enable Automatic Updates   ------- ####
  # Set-ODTUpdates -TargetFilePath $targetFilePath -Enabled $true -Channel $Channel


  #### ------- Disable Automatic Updates   ------- ####
  # Set-ODTUpdates -TargetFilePath $targetFilePath -Enabled $false
 
 #------------------------------------------------------------------------------------------------------------

 #Add excluded apps
 if($ExcludedApps){
    if($ExcludedApps -match ","){
        $ExcludedApps = $ExcludedApps.Split(",")
    }
    foreach($app in $ExcludedApps){
        Exclude-Applications -TargetFilePath $targetFilePath -ExcludeApps $app
    }
 }

 #Add additional apps
 if($AdditionalApps){
    if($AdditionalApps -match ","){
        $AdditionalApps = $AdditionalApps.Split(",")
    }
    foreach($app in $AdditionalApps){
        Add-ProductSku -TargetFilePath $targetFilePath -Languages $languages -ProductIDs $app
    }
 }

 #Add languages to each product
 Add-ProductLanguage -TargetFilePath $targetFilePath -ProductIDs All -Language $Languages

 # Installs Office 365 ProPlus
 Install-OfficeClickToRun -TargetFilePath $targetFilePath
 
 # Configuration.xml file for Click-to-Run for Office 365 products reference. https://technet.microsoft.com/en-us/library/JJ219426.aspx
}


