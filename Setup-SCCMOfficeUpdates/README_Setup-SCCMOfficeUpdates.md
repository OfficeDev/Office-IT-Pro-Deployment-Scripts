##**Update Office 2013 using SCCM**

This script automates the configuration updating of Office 365 ProPlus through Microsoft System Center Configuration Manager (SCCM) and ensures the PC where Office 365 ProPlus is installed gets Office 365 ProPlus updates from the closest SCCM Distribution Point (DP).

###**Pre-Requisites:**

Before running this script, the following conditions must be met

1. .Net Framework 3.5 SP1 must be installed on client machines.
2. An operational SCCM environment.
3. Office 2013 or Office 2016 is already installed on client machines. 
4. Office Auto Updates have been Disabled on the client machines preferably via Group Policy.

###**Running the script**

1. The script should ideally be run on a SCCM server. If you do not run it from a SCCM server ensure you always specify the -Path parameter.
2. Open an Elevated PowerShell Console(see, [Starting Windows PowerShell](https://technet.microsoft.com/en-us/library/hh857343.aspx)):

		From the Run dialog type PowerShell.

3. Change directory to the location where the PowerShell Script is saved. This directory must contain the files that are in the *Setup-SCCMOfficeUpdates* folder.

		Example: cd C:\PowerShellScripts

4. Type the following in the elevated PowerShell Session

		 . .\Setup-SCCMOfficeUpdates.ps1
         
         By including the additional period before the relative script path you are 'Dot-Sourcing' 
		 the PowerShell functions in the script into your PowerShell session which will allow you to 
		 run the function from the console.

5. Before you download the Updates you must configure the following files with the products and languages that you have installed in your environment.  This will ensure that you have the files necessary to update the products in your environment.  For reference of the options in the configuration xml you can either go to https://technet.microsoft.com/en-us/library/JJ219426.aspx or we provide an online editor in this GitHub repository at http://officedev.github.io/Office-IT-Pro-Deployment-Scripts/XmlEditor.html

	If you specify a version in either the xml files below or in the command line options, in order to make that version the active version for this solution you either have to copy the v32_XX.X.XXXX.XXXX.cab and v64_XX.X.XXXX.XXXX.cab and rename them to v32.cab and v64.cab respectively. If no version is specified, the latest version will be downloaded and the v32.cab and v64.cab files will be updated automatically.

		Configuration_UpdateSource32.xml
		Configuration_UpdateSource64.xml

6. The first thing you must do is download the Office update files to a staging location to make them available for SCCM. From the existing PowerShell session type the command below.

		Download-OfficeUpdates -Path (Optional) -Version (Optional)
        
	If you specify the *-Path* parameter then the script will download the Office updates to that path. The path must be a valid UNC path. Specifying the *-Version* parameter will cause the script to download a specific version of the Office updates.
    
    If you do not specify any parameters the script will create a local folder name 'OfficeUpdates' on the SystemDrive.  It will then share the folder using a hidden share name 'OfficeUpdates$'. This share will be used to store the Office update files. If you are not running the script on a SCCM server it is important that you specify the -Path parameter with all functions so the local share will not be created.
    
7. Now that the Office update files have been downloaded to a share on the network you can run the setup function to configure SCCM. A SCCM collection must be specified to use this function. The collection specified should contain the workstations that you want configured.  If there are no Distribution Point Groups added to the collection then you will also have to use the parameter *-DistributionPointGroupName*

		Setup-SCCMOfficeUpdates -Collection CollectionName -DistributionPointGroupName DPGroupName

8. The function *Setup-SCCMOfficeUpdates* will create a SCCM Package that is configured to run the *SCO365PPTrigger.exe* executable on the client machines.  After the package is created the cmdlet *Start-CMContentDistribution* is run in order to start the process to distribute the Package contents to the Distribution Points. Before proceeding to the next step you should monitor and wait until the content distribution process is complete.  Clients deployments will fail until the content is distributed to their Distribution Point.

9. Once the content distribution is complete you can run the final function.  The *Deploy-SCCMOfficeUpdates* function will deploy the package.  You must provide the name of the collection.  This should be the same collection you used with the previous function.

		Deploy-SCCMOfficeUpdates -Collection CollectionName

