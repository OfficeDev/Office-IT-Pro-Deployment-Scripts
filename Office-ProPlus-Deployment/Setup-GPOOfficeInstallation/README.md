### Configure GPO Office Installation
This script will configure an existing Active Directory Group Policy to silently install Office 2013 Click-To-Run on computer startup.

**IT Pro Scenario:** For organizations that aren't using SCCM or another managed deployment system, this script will install Office using a Group Policy. 

[README](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/wiki/README_New-GPOOfficeInstallation)

#### Download the Office channel files
1. Open an elevated PowerShell console.

		From the Run dialog type PowerShell and run as administrator

2. Change directory to the location where the PowerShell Script is saved on your local machine.

		Example: cd C:\Setup-GPOOfficeInstallation

3. Dot-Source the Configure-GPOOfficeInstallation function into your current session.

		Type . .\Setup-GPOOfficeInstallation.ps1
		
		By including the additional period before the relative script path you are 'Dot-Sourcing' 
		the PowerShell function in the script into your PowerShell session which will allow you to 
		run the function from the console.

4. Download the Office channel files you plan to deploy in your environment

		For example, type Download-GPOOfficeChannelFiles -Channels Deferred,FirstReleaseDeferred -OfficeFilesPath C:\OfficeChannelFiles -Languages en-us,de-de -Bitness v32

		In this example, the latest 32-bit versions of the Deferred and FirstReleaseDeferred 
		channels will be downloaded to C:\OfficeChannelFiles. If C:\OfficeChannelFiles does not 
		exist a new directory will be created. Both English and German languages will be downlaoded.


#### Configure the OfficeDeployment folder used to stage the Office channel files and the PowerShell scripts
1. Create the OfficeDeployment$ folder.

		Type Configure-GPOOfficeDeployment -Channel Deferred,FirstReleaseDeferred -Bitness v32 -OfficeFilesPath C:\OfficeChannelFiles -MoveSourceFiles $true

		A new directory will be created on your largest drive called OfficeDeployment$ and 
		all of the necessary files will be copied here. 


#### Create a new Office deployment using an existing GPO
1. Install Office using dynamic PowerShell scripts

		Type Create-GPOOfficeDeployment -GroupPolicyName "DeployDeferredChannel" -DeploymentType DeployWithScript -Channel Deferred -Bitness v32

2. Install Office using a custom configuration.xml file

		Type Create-GPOOfficeDeployment -GroupPolicyName "DeployDeferredChannel" -DeploymentType DeployWithConfigurationFile -Channel Deferred -Bitness v32 -ConfigurationXML Deferred-Channel-Configuration.xml

		Make sure you've saved the configuration xml files into the OfficeDeployment$ folder.

3. Install Office using a packaged MSI or Executable file

		Type Create-GPOOfficeDeployment -GroupPolicyName "DeployDeferredChannel" -DeploymentType DeployWithInstallationFile -OfficeDeploymentFileName OfficeProPlus.msi -Quiet $true

		This is example is intended to use an installation MSI or executable file that was generated using 
		the Office 365 ProPlus Toolkit that can located from 
		the [Office Click-To-Run Configuration XML Editor](http://officedev.github.io/Office-IT-Pro-Deployment-Scripts/XmlEditor.html)

#### Remove previous versions of Office
1. Remove the previous versions of Office

		Type Create-GPOOfficeDeployment -GroupPolicyName "DeployDeferredChannel" -DeploymentType RemoveWithScript

		This example will remove all previous versions of Office using the 
		GPO-ExampleRemovePreviousOfficeInstalls.ps1 script.

2. Remove the previous versions of Office with a custom script

		Type Create-GPOOfficeDeployment -GroupPolicyName "DeployDeferredChannel" -DeploymentType RemoveWithScript -ScriptName GPO-RemovePreviousOfficeInstalls.ps1

		

		





