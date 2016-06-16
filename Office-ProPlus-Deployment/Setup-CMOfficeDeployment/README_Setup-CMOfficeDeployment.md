#Setup Config Manager Office Deployment

This PowerShell function automates the setup of Office 365 Click-To-Run deployment and update scenarios in Config Manager. 

[README](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/wiki/Readme_Setup-CMOfficeDeployment)

##Scenarios for creating, updating, and deploying the Office 365 ProPlus programs.
###Deploy Office 365 ProPlus
1. Download the channel files from the CDN.

		Download-CMOOfficeChannelFiles -Channels Deferred -Bitness v32 -OfficeFilesPath D:\OfficeChannels
		
	* The Deferred channel 32 bit files will be downloaded to D:\OfficeChannels.

2. Create the Office 365 ProPlus package.

		Create-CMOfficePackage -Channels Deferred -OfficeSourceFilesPath D:\OfficeChannels -MoveSourceFiles $true -SiteCode S01 -Bitness v32
		
	* A package will be created called Office 365 ProPlus. The source files will be moved from D:\OfficeChannels to a new folder called OfficeDeployment$.

3. Create the deployment program.
 
		Create-CMOfficeDeploymentProgram -Channels Deferred -Bitness v32 -DeploymentType DeployWithConfigurationFile -SiteCode S01
		
	* A program called "Deploy Deferred Channel with Config File - 32-Bit" will be created.

4. Distribute the package to the distribution point.
 
		Distribute-CMOfficePackage -Channels Deferred -DistributionPoint cm.contoso.com
		
	* The files in the OfficeDeployment$ folder will be distributed to the distribution point called cm.contoso.com.

5. Deploy the program to the collection.
 
		Deploy-CMOfficeProgram -Collection 'Human Resources' -Channel Deferred -Bitness v32 -SiteCode S01 -DeploymentPurpose Available
		
	* A deployment will be created and made available to the collection 'Human Resources' that will install the "Deploy Deferred Channel with Config File - 32-Bit" program.

###Change the channel of an Office 365 client.
1. Download the channel files from the CDN.

		Download-CMOOfficeChannelFiles -Channels FirstReleaseDeferred -Bitness v32 -OfficeFilesPath D:\OfficeChannels
		
	* The FirstReleaseDeferred 32 bit channel files will be downloaded to D:\OfficeChannels.

2. Create the Office 365 ProPlus package.

		Create-CMOfficePackage -Channels Deferred -OfficeSourceFilesPath D:\OfficeChannels -MoveSourceFiles $true -SiteCode S01 -Bitness v32
		
	* A package will be created called Office 365 ProPlus. The source files will be moved from D:\OfficeChannels to a new folder called OfficeDeployment$.
		
	**Note** - You do not need to create the Office 365 ProPlus package if it already exists.
		
3. Create the channel change program.

		Create-CMOfficeChannelChangeProgram -Channels FirstReleaseDeferred

4. Update the package with the new channel files.
	
		Update-CMOfficePackage -Channels FirstReleaseDeferred -OfficeSourceFilesPath D:\OfficeChannels -MoveSourceFiles $true

	* The FirstReleaseDeferred channel files will be moved to OfficeDeployment$ and the cm.contoso.com distribution point will be updated.

5. Create a deployment.

		Deploy-CMOfficeProgram -Collection 'Human Resources' -ProgramType DeployWithConfigurationFile -Channel FirstReleaseDeferred -Bitness v32 -SiteCode S01 -DeploymentPurpose Available

###Rollback the version of an Office 365 client.

1. Download the channel version files.

		Download-CMOOfficeChannelFiles -Channels FirstReleaseDeferred -Version 16.0.6741.2042 -Bitness v32 -OfficeFilesPath D:\OfficeChannels

2. Create the Office 365 ProPlus package.

		Create-CMOfficePackage -Channels Deferred -OfficeSourceFilesPath D:\OfficeChannels -MoveSourceFiles $true -SiteCode S01 -Bitness v32
		
	* A package will be created called Office 365 ProPlus. The source files will be moved from D:\OfficeChannels to a new folder called OfficeDeployment$.
		
	**Note** - You do not need to create the Office 365 ProPlus package if it already exists.

3. Create the rollback program.

		Create-CMOfficeRollBackProgram

4. Update the package with the new channel version files.

		Update-CMOfficePackage -Channels FirstReleaseDeferred -OfficeSourceFilesPath D:\OfficeChannels -MoveSourceFiles $true

5. Create a deployment.

		Deploy-CMOfficeProgram -Collection 'Human Resources' -ProgramType RollBack -Channel FirstReleaseDeferred -Bitness v32 -SiteCode S01 -DeploymentPurpose Available

###Update an Office 365 ProPlus client with ConfigMgr

1. Download the channel files to update the client.

		Download-CMOOfficeChannelFiles -Channels FirstReleaseDeferred -Bitness v32 -OfficeFilesPath D:\OfficeChannels

2. Create the Office 365 ProPlus package.

		Create-CMOfficePackage -Channels Deferred -OfficeSourceFilesPath D:\OfficeChannels -MoveSourceFiles $true -SiteCode S01 -Bitness v32
		
	* A package will be created called Office 365 ProPlus. The source files will be moved from D:\OfficeChannels to a new folder called OfficeDeployment$.
		
	**Note** - You do not need to create the Office 365 ProPlus package if it already exists.

3. Create the update program.

		Create-CMOfficeUpdateProgram -WaitForUpdateToFinish $true -EnableUpdateAnywhere $true -ForceAppShutdown $true -SiteCode S01 -UpdateToVersion 16.0.6741.2047 	

4. Update the package with the new channel version files.

		Update-CMOfficePackage -Channels FirstReleaseDeferred -OfficeSourceFilesPath D:\OfficeChannels -MoveSourceFiles $true

5. Create a deployment.

		Deploy-CMOfficeProgram -Collection 'Human Resources' -ProgramType UpdateWithConfigMgr -Channel FirstReleaseDeferred -Bitness v32 -SiteCode S01 -DeploymentPurpose Required

###Update an Office 365 ProPlus client using a scheduled task.

1. Download-CMOOfficeChannelFiles

		Download-CMOOfficeChannelFiles -Channels FirstReleaseDeferred -Bitness v32 -OfficeFilesPath D:\OfficeChannels

2. Create the Office 365 ProPlus package.

		Create-CMOfficePackage -Channels Deferred -OfficeSourceFilesPath D:\OfficeChannels -MoveSourceFiles $true -SiteCode S01 -Bitness v32
		
	* A package will be created called Office 365 ProPlus. The source files will be moved from D:\OfficeChannels to a new folder called OfficeDeployment$.
		
	**Note** - You do not need to create the Office 365 ProPlus package if it already exists.

3. Create-CMOfficeUpdateAsTaskProgram

		Create-CMOfficeUpdateAsTaskProgram -WaitForUpdateToFinish $true -EnableUpdateAnywhere $true -ForceAppShutdown $true -SiteCode S01 -UpdateToVersion 16.0.6741.2047
		
4. Update the package with the new channel version files.

		Update-CMOfficePackage -Channels FirstReleaseDeferred -OfficeSourceFilesPath D:\OfficeChannels -MoveSourceFiles $true

5. Create a deployment.

		Deploy-CMOfficeProgram -Collection 'Human Resources' -ProgramType UpdateWithTask -Channel FirstReleaseDeferred -Bitness v32 -SiteCode S01 -DeploymentPurpose Required

##Creating the Office ProPlus Package
###Prepare the environment

1. Download the **Setup-CMOfficeDeployment** folder to your Config Manager Server. 
2. Open PowerShell as an administrator.

		From the Run dialog type PowerShell, right click it and choose Run as Administrator.
		
3. Change the directory to the location where the PowerShell Script is saved. 

		Example: cd C:\PowerShellScripts
		
4. Dot-Source the script to gain access to the functions inside.

		Type: . .\Setup-CMOfficeDeployment.ps1
		
		By including the additional period before the relative script path you are 'Dot-Sourcing' 
   		the PowerShell function in the script into your PowerShell session which will allow you to 
   		run the inner functions from the console.

###Downloading the source files from the CDN

1. Run **Download-CMOOfficeChannelFiles**. This function will download all the source files from the CDN.

	The available parameters with the function are as follows.
	* **Channels** The available options are **Current, Deferred, FirstReleaseDeferred, FirstReleaseCurrent** If it is not specified all channels will be downloaded. 
	* **OfficeFilesPath** The location to download the source files.
	* **Languages** Uses the ll-cc office codes found [Here](https://technet.microsoft.com/en-us/library/cc179219.aspx) 
	* **Bitness**  Available options are **v32, v64, Both**. Default value is Both.
	* **Version** You can specify a version to download. Versions and the associated channels can be found [Here](https://technet.microsoft.com/en-us/library/mt592918.aspx)
	
			Example: Download-CMOfficeChannels -Channels Deferred,FirstReleaseDeferred -OfficeFilesPath D:\OfficeChannels -Languages en-us,es-es,de-de -Bitness v32

###Create the Office ProPlus package

1. Run **Create-CMOfficePackage**. This function creates the package and associated package share

	The available parameters with the function are as follows.
	* **Channels** The available options are **Current, Deferred, FirstReleaseDeferred, FirstReleaseCurrent**
	* **OfficeSourceFilesPath** The location the source files are located
	* **MoveSourceFiles** Moves the source files to the new package share vs copying
	* **CustomPackageShareName** Create a custom package share name. Default value is OfficeDeployment.
	* **UpdateOnlyChangedBits** Used when the package share already exists.
	* **SiteCode** Three digit site code, example **S01**. Left blank it will default to the current site. 
	* **Bitness** Available options are **v32, v64, Both**. Default value is Both.
	* **CMPSModulePath** Allows the user to specify that full path to the ConfigurationManager.psd1 PowerShell Module. This is especially useful if CM is installed in a non standard path.
	
			Example: Create-CMOfficePackage -Channels Deferred -OfficeSourceFilesPath D:\OfficeChanels -MoveSourceFiles $true -SiteCode S01 -Bitness v32

##Updating the Office 365 ProPlus package

1. To update the Office ProPlus package use **Update-CMOfficePackage**

	The available parameters with the function are as follows.
	* **Channels** The available options are **Current, Deferred, FirstReleaseDeferred, FirstReleaseCurrent** 
	* **OfficeSourceFilesPath** The location the source files are located
	* **MoveSourceFiles** Moves the source files to the new package share vs copying
	
			Example: Update-CMOfficePackage -Channels FirstReleaseDeferred -OfficeSourceFilesPath D:\OfficeChannels -MoveSourceFiles $true

##Creating Office 365 Client Programs
###Create-CMOfficeDeploymentProgram
1. To create an Office 365 deployment program use **Create-CMOfficeDeploymentProgram**

	The available parameters with the function are as follows.
	* **Channels** The available options are **Current, Deferred, FirstReleaseDeferred, FirstReleaseCurrent** 
	* **Bitness** Available options are **v32, v64, Both**. Default value is Both.
	* **DeploymentType** The available options are **DeployWithScript,DeployWithConfigurationFile**
	* **ScriptName** Default value is **CM-OfficeDeploymentScript.ps1**
	* **SiteCode** Three digit site code, example **S01**. Left blank it will default to the current site. 
	* **CMPSModulePath** Default value will use the default location.
	* **ConfigurationXml** Default value is **.\DeploymentFiles\DefaultConfiguration.xml**
	* **CustomName** Default value combines the channel with the platform.
	
			Example: Create-CMOfficeDeploymentProgram -Channels Deferred,FirstReleaseDeferred -Bitness v32 -DeploymentType DeployWithConfigurationFile -SiteCode S01

###Create-CMOfficeChannelChangeProgram

1. To create an Office 365 channel change program use **Create-CMOfficeChannelChangeProgram**
	
	The available parameters with the function are as follows.
	* **Channels** The available options are **Current, Deferred, FirstReleaseDeferred, FirstReleaseCurrent**
	* **SiteCode** Three digit site code, example **S01**. Left blank it will default to the current site.
	* **CMPSModulePath** Default value will use the default location.

			Example: Create-CMOfficeChannelChangeProgram -Channels Deferred -SiteCode S01
			
###Create-CMOfficeRollBackProgram

1. To create an Office 365 rollback program use **Create-CMOfficeRollBackProgram**
	
	The available parameters with the function are as follows.
	* **SiteCode** Three digit site code, example **S01**. Left blank it will default to the current site.
	* **CMPSModulePath** Default value will use the default location.

			Example: Create-CMOfficeRollBackProgram
			
###Create-CMOfficeUpdateProgram

1. To create an Office 365 client update program use **Create-CMOfficeUpdateProgram**
	
	The available parameters with the function are as follows.
	* **WaitForUpdateToFinish** The PowerShell window will remain open until the update has finished. Default value is $true.
	* **EnableUpdateAnywhere** The failback method if the update path is unavailable the client will update from the CDN. Default value is $true.
	* **ForceAppShutdown** Default value is $false.
	* **UpdatePromptUser** Default value is $false.
	* **DisplayLevel** Default value is $false.
	* **UpdateToVersion** The version to update to. Default value will update to the latest version in the update path.
	* **LogPath** The path to the LogName.
	* **LogName** The name of the log files.
	* **ValidateUpdateSourceFiles** Default value is $true.
	* **SiteCode** Three digit site code, example **S01**. Left blank it will default to the current site.
	* **CMPSModulePath** Default value will use the default location.
	* **UseScriptLocationAsUpdateSource** If not specified the location where the script is ran will be assumed the location of the SourceFiles. Default value is $true.

			Example: Create-CMOfficeUpdateProgram -WaitForUpdateToFinish $true -EnableUpdateAnywhere $true -ForceAppShutdown $true -SiteCode S01
			
###Create-CMOfficeUpdateAsTaskProgram

1. To create an Office 365 update program that will run as a task use **Create-CMOfficeUpdateAsTaskProgram**
	
	The available parameters with the function are as follows.
	* **WaitForUpdateToFinish** The PowerShell window will remain open until the update has finished. Default value is $true.
	* **EnableUpdateAnywhere** The failback method if the update path is unavailable the client will update from the CDN. Default value is $true.
	* **ForceAppShutdown** Default value is $false.
	* **UpdatePromptUser** Default value is $false.
	* **DisplayLevel** Default value is $false.
	* **UpdateToVersion** The version to update to. Default value will update to the latest version in the update path.
	* **UseRandomStartTime** Default value is $true.
	* **RandomTimeStart** Default value is 08:00.
	* **RandomTimeEnd** Default value is 17:00.
	* **StartTime** Default value 12:00.
	* **LogPath** The path to the LogName.
	* **LogName** The name of the log files.
	* **ValidateUpdateSourceFiles** Default value is $true.
	* **SiteCode** Three digit site code, example **S01**. Left blank it will default to the current site.
	* **CMPSModulePath** Default value will use the default location.
	* **UseScriptLocationAsUpdateSource** If not specified the location where the script is ran will be assumed the location of the SourceFiles. Default value is $true.

			Example: Create-CMOfficeUpdateAsTaskProgram -WaitForUpdateToFinish $false -EnableUpdateAnywhere $false -ForceAppShutdown $true -UpdatePromptUser $true -UpdateToVersion 16.0.6001.1078
			
##Distribute the Office 365 ProPlus Package

1. To distribute the Office 365 package use **Distribute-CMOfficePackage**
	
	The available parameters with the function are as follows.
	* **Channels** The available options are **Current, Deferred, FirstReleaseDeferred, FirstReleaseCurrent**
	* **DistributionPoint** The distribution point name. A distribution point or distirbution point group must be specified.
	* **DistributionPointGroupName** The distribution point group name. A distribution point or distirbution point group must be specified.
	* **DeploymentExpiryDurationInDays** Default value is 15.
	* **SiteCode** Three digit site code, example **S01**. Left blank it will default to the current site.
	* **CMPSModulePath** Default value will use the default location.

			Example: Distribute-CMOfficePackage -Channels Deferred -DistributionPoint cm.contoso.com
			
##Deploy the Office 365 ProPlus programs

1. To create an Office 365 deployment use **Deploy-CMOfficeProgram**

	The available parameters with the function are as follows.
	* **Collection** The name of the collection to deploy the program to.
	* **ProgramType** The type of program being deployed. Available options are **DeployWithScript,DeployWithConfigurationFile,ChangeChannel,RollBack,UpdateWithConfigMgr,UpdateWithTask** 
	* **Channel** The available options are **Current, Deferred, FirstReleaseDeferred, FirstReleaseCurrent**
	* **Bitness** Available options are **v32, v64, Both**. Default value is Both.
	* **SiteCode** Three digit site code, example **S01**. Left blank it will default to the current site.
	* **CMPSModulePath** Default value will use the default location.
	* **DeploymentPurpose** Default value is Required. Available options are **Default,Required,Available**
	* **CustomName** Default value combines the channel with the platform.

			Example: Deploy-CMOfficeProgram -Collection 'Human Resources' -ProgramType DeployWithConfigurationFile -Channel Deferred -Bitness v32 -SiteCode S01 -DeploymentPurpose Available
