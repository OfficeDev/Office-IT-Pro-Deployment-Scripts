cls
$RGName = "RG-o365-W10"; 
$VMName = "jdwino365";
$VMUsername = "jmd";

$ARMTemplate = "C:\Dump\O365 Testing\azuredeploy.json"
$DeployLocation = "West Europe"
$OfficeVersion  = "Office2013"; #or Office2016

# 1. Login
#Login-AzureRmAccount

#2. Create a resource group
New-AzureRmResourceGroup -Name $RGName -Location $DeployLocation -Force

#3. Create resources
$sw = [system.diagnostics.stopwatch]::startNew()
New-AzureRmResourceGroupDeployment -ResourceGroupName $RGName -TemplateFile $ARMTemplate -vmName $VMName `
    -vmAdminUserName $VMUsername -dnsLabelPrefix $VMName -vmVisualStudioVersion VS-2015-Comm-AzureSDK-2.9-W10T-Win10-N `
    -officeVersion "Office2013" -Mode Complete -Force | Out-Null
$sw | Format-List -Property *

#4. Get the RDP file
Get-AzureRmRemoteDesktopFile -ResourceGroupName $RGName -Name $VMName -Launch -Verbose -Debug

