# Get Script Directory
if ($PSScriptRoot) {
  $scriptPath = $PSScriptRoot
} else {
  $scriptPath = (Get-Item -Path ".\").FullName
}

# Importing the required functions
. $scriptPath\Download-OfficeProPlusChannels.ps1 
. $scriptpath\Generate-ODTConfigurationXML.ps1
. $scriptPath\Edit-OfficeConfigurationFile.ps1

# Set the file path parameters
$TargetFilePath = "$scriptPath\O365"
$DefaultConfigurationFile = "$scriptPath\DefaultConfiguration.xml"

# Set the channels, bit, and languages to download
$Channels = @("Current","Deferred","FirstReleaseDeferred")
$Bitness = @("v32")
$Languages = @("en-us","de-de")

# Get ODT Download URI
$odtUri = ((Invoke-WebRequest -Uri 'https://www.microsoft.com/en-us/download/confirmation.aspx?id=49117').Links | Where-Object {$_.href -like "*officedeploymenttool*"}).href

# Get ODT File Name from URI
$odtFileName = $odtUri.Substring($odtUri.LastIndexOf("/") + 1)

# Download Latest ODT
Start-BitsTransfer -Source $odtUri -Destination $scriptPath

# Build ODT File Path
$odtFilePath = Join-Path -Path $scriptPath -ChildPath $odtFileName

# Extract ODT from compressed Download
& $odtFilePath /extract:$TargetFilePath /quiet /norestart

# Download the channel files
Download-OfficeProPlusChannels -Channels $Channels -TargetDirectory $TargetFilePath -Bitness v32 -Languages $Languages

# Generate the ODT configuration files
foreach($channel in $Channels){
    # Create the configuration file
    $configFileName = "Deploy-$channel-$Bitness"
    $path = "$TargetFilePath\$configFileName" + ".xml"
    Copy-Item $DefaultConfigurationFile $path

    # Set the SourcePath
    $ChannelShortName = ConvertChannelNameToShortName -ChannelName $channel
    $SourcePath = "$TargetFilePath\$ChannelShortName"
    Set-ODTAdd -TargetFilePath $path -SourcePath $SourcePath | Out-Null

    #------------------------------------------------------------------------------------------------------------
    #   Customize Deployment Script - Uncomment and modify the code below to customize this deployment script
    #------------------------------------------------------------------------------------------------------------
   
    #### ------- Exclude Applications ------- ####
    # Exclude-Applications -TargetFilePath $path -ExcludeApps @("Access","Excel","Groove","InfoPath","Lync","OneDrive","OneNote","Outlook","PowerPoint","Project","Publisher","SharePointDesigner","Visio","Word") | Out-Null
 
    #### ------- Add an additional Product Sku ------- ####
    # Add-ProductSku -TargetFilePath $path -Languages $languages -ProductIDs O365ProPlusRetail,O365BusinessRetail,VisioProRetail,ProjectProRetail | Out-Null

    #### ------- Set the display to Full so the installation   ------- ####
    # Set-ODTDisplay -TargetFilePath $path -Level Full -AcceptEULA $true | Out-Null

    #### ------- Enable Automatic Updates   ------- ####
    # Set-ODTUpdates -TargetFilePath $path -Enabled $true -Channel $Channel | Out-Null

    #### -------- Enable logging ------- ####
    # Set-ODTLogging -TargetFilePath $path -Level Standard | Out-Null

    #### -------- Set additional properties ------- ####
    # Set-ODTConfigProperties -TargetFilePath $path -AutoActivate $true -ForceAppShutDown $true -SharedComputerLicensing $true -PinIconsToTaskbar $false | Out-Null

    #------------------------------------------------------------------------------------------------------------
}