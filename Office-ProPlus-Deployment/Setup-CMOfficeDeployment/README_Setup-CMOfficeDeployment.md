# Setup Config Manager Office Deployment

This PowerShell function automates the setup of Office 365 Click-To-Run deployment and update scenarios in Config Manager. 

[README](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/wiki/Readme_Setup-CMOfficeDeployment)

## Creating the Office ProPlus Package

###Prepare the environment

1. Download the **Setup-CMOfficeDeployment** folder to your Config Manager Server. 
2. Load an elevated Powershell session
3. Change the directory to the **Setup-CMOfficeDeployment** location
4. From the prompt run **. .\Setup-CMOfficeDeployment.ps1** this dot sources the file and enables all functions and variables to be available in the power shell session.

###Downloading the source files from the CDN

1. Run **Download-CMOOfficeChannelFiles** this function will download all the source files from the CDN.
     The available parameters with the function are as follows.
	* a. **-Channels** This parameter defines what channels to download. If it is not specified all channels will be downloaded. The available options are **Current, Deferred, FirstReleaseDeferred, FirstReleaseCurrent**
	* b. **-OfficeFilesPath** The location to download the source files into.
	* c. **-Languages** Uses the ll-cc office codes found [Here](https://technet.microsoft.com/en-us/library/cc179219.aspx) 
	* d. **-Bitness**  Left blank it will download both versions into source. Options are **v32, v64, Both**
	* e. **-Version** You can specify a version to download. Versions and the associated channels can be found [Here](https://technet.microsoft.com/en-us/library/mt592918.aspx)

###Create the Office ProPlus Package

1. Run **Create-CMOfficePackage** this function creates the package and associated package share
     The available parameters with the function are as follows.
	* a. **Channels** The available options are **Current, Deferred, FirstReleaseDeferred, FirstReleaseCurrent**
	* b. **OfficeSourceFilesPath** The location the source files are located
	* c. **MoveSourceFiles** Moves the source files to the new package share vs copying
	* d. **CustomPackageShareName** You can define the name of the package share. Left blank it will default to OfficeDeployment
	* e. **UpdateOnlyChangedBits** Used when the package share already exists.
	* f. **SiteCode** Three digit site code, example **S01**. Left blank it will default to the current site. 
	* g. **Bitness** Left blank it will create a package with v32. Options are **v32, v64, Both**
	* h. **CMPSModulePath** Allows the user to specify that full path to the ConfigurationManager.psd1 PowerShell Module. This is especially useful if CM is installed in a non standard path.

###Updating the Office ProPlus Package

1. To update the Office ProPlus package use **Update-CMOfficePackage**
     The available parameters with the function are as follows.
	* a. **Channels** The available options are **Current, Deferred, FirstReleaseDeferred, FirstReleaseCurrent** 
	* b. **OfficeSourceFilesPath** The location the source files are located
	* c. **MoveSourceFiles** Moves the source files to the new package share vs copying
	




	Create-CMOfficeDeploymentProgram
	Create-CMOfficeChannelChangeProgram
	Create-CMOfficeRollBackProgram
	Create-CMOfficeUpdateProgram
	Create-CMOfficeUpdateAsTaskProgram
Distribute-CMOfficePackage
Deploy-CMOfficeProgram

Scenaro: Install Office

1) Download-CMOOfficeChannelFiles
2) Create-CMOfficePackage
3) Create-CMOfficeDeploymentProgram
4) Distribute-CMOfficePackage
5) Deploy-CMOfficeProgram

Scenario: Channel Change

1) Download-CMOOfficeChannelFiles
2) Create-CMOfficePackage
3) Create-CMOfficeChannelChangeProgram
4) Distribute-CMOfficePackage
5) Deploy-CMOfficeProgram

Scenario: Rollback

For roll back you need to have the version in source to roll back to.

1) Download-CMOOfficeChannelFiles
2) Create-CMOfficePackage
3) Create-CMOfficeRollBackProgram
4) Distribute-CMOfficePackage
5) Deploy-CMOfficeProgram

Scenario: Update Office

1) Download-CMOOfficeChannelFiles
2) Create-CMOfficePackage
3) Create-CMOfficeUpdateProgram
4) Distribute-CMOfficePackage
5) Deploy-CMOfficeProgram

Scenario: Update Office

1) Download-CMOOfficeChannelFiles
2) Create-CMOfficePackage
3) Create-CMOfficeUpdateAsTaskProgram
4) Distribute-CMOfficePackage
5) Deploy-CMOfficeProgram

