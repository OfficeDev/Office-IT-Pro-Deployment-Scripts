#**Install or Upgrade Office 365 ProPlus**

This guide provides the different options for installing or upgrading Office 365 ProPlus. We've divided the guide into 3 sections:  

Unmanaged deployments - For organizations that are not using managed deployment software like System Center Configuration Manager and Microsoft Intune. 
Several examples are shown on how to deploy Office ProPlus using PowerShell.  

Managed deployments - Deploy Office 365 ProPlus through System Center Configuration Manager using PowerShell commands. In addition, instructions for deploying
Office ProPlus through Group Policy are also outlined.

Package Office deployments into an MSI or EXE - Package an Office 365 ProPlus installation into an MSI or Executable file using the Microsoft Office ProPlus Install Toolkit.  

##**Before you run a script**
1. Copy the folder containing the necessary PowerShell scripts and files locally.  
2. Run PowerShell with elevated privileges.  
	a. In Windows 10, from the Cortana search box type PowerShell, right click on Windows PowerShell and choose Run as Administrator.
	If you need to enter the credentials of a privileged account, right click Windows PowerShell and choose Pin to Taskbar.
	Hold the Shift key and right click on the Windows PowerShell icon and choose Run as different user.  
	b. In Windows 8.1, go to the Start Menu and type powershell. right click on Windows PowerShell and choose Run as Administrator.
	If you need to enter the credentials of a privileged account, right click Windows PowerShell and choose Pin to Taskbar, and open the Desktop.
	Hold the Shift key and right click on the Windows PowerShell icon, and choose Run as different user.  
	c. In Windows 7, click the Start button and type powershell. Right click on Windows PowerShell and choose Run as administrator.
	If you need to enter the credentials of a privileged account hold the Shift key and right click on Windows PowerShell and choose Run as different user.  
3. Change the directory to the Deploy-OfficeClickToRun folder.  
	a. For example, type **cd C:\Deploy-OfficeClickToRun and press Enter.**  
4. Set the execution policy to unrestricted or bypass.  
	a. In the PowerShell console, type **Set-ExecutionPolicy Unrestricted**, press **Enter**, and accept the policy change request. 
	For more information about execution policies please visit http://go.microsoft.com/fwlink/?LinkID=135170.  
Note: By typing **.\** before a script name lets PowerShell know to execute a script located in the current directory. 
PowerShell also uses IntelliSense which will allow you to start typing the name of a script, cmdlet, or function and press tab to finish the name.  

##**Use PowerShell to install or upgrade to Office ProPlus**
###Scenario: Install Office ProPlus and keep all in use language
This script generates an Office Deployment Tool (ODT) configuration.xml file and will include in use languages on the system.  
1. From the PowerShell console type **.\1-ExampleDeployGeneric.ps1** and press **Enter**.  

###Scenario: Install Office 365 ProPlus silently 
This script is similar to 1-ExampleDeployGeneric.ps1, but the configuration.xml will be modified to set the DisplayLevel to None.  
1. From the PowerShell console type **.\5-ExampleDeploySilent.ps1** and press **Enter**.  

###Scenario: Install Office ProPlus using OU filters
This script may be particularly useful if you have locations within the organization that only need a specific language installed. 
A configuration.xml file will be generated, then the computer will be checked against a specified OU and install the language for that region. 
We have included 2 examples for OUs named Paris and Tokyo.  
1. From the directory where the scripts are saved right click on **2-ExampleDeployWithOfficeFilter.ps1** and click **Edit**. 
Windows PowerShell ISE will open and we will be able to modify the script before we run it.   
2. Replace the example OU with the OU that hosts the computers of the region. If your OU is named "North America" you would replace "OU=Paris" with "OU=North America".  
3. If you want Office to update from a particular file share replace the UpdatePath with the respective file share. 
If the file share for the computers in the North America OU is named "NorthAmerica\Share" replace "\\ParisFileServer\OfficeUpdates" with \\NorthAmerica\Share.   
4. Copy and replace the OU examples as many times as needed. A new script block will be needed for each OU. 
This is why we have separate script blocks for Paris and Tokyo.  
5. When you are finished editing the script click **File** and choose **Save**.   
6. Close the Windows PowerShell ISE window.  
7. From the Windows PowerShell console type **.\2-ExampleDeployWithOfficeFilter.ps1** and press **Enter**.  

###Scenario: Install Office ProPlus using a dynamic source path 
This script compares the site of the computer to be upgraded with SourcePathLookup.csv. 
SourcePathLookup.csv must be updated first to reflect the sites and local source of the Office ProPlus install files.  
1. Open SourcePathLookup.csv.  
2. Replace Site1, Site2, and Site3 and their corresponding Sources to your organizations Sites and Source paths.   
3. Save and close the document.  
4. From the Windows PowerShell console type **.\6-ExampleDeployWithDynamicSourcePath.ps1** and press **Enter**.  

###Scenario: Install Office ProPlus with additional Office ProPlus products 
This script will install Office ProPlus and any additional Office ProPlus products of your choice. 
There are 2 examples in the script to show how to include Visio and Project with the Office ProPlus deployment.  
1. From the directory where the scripts are saved right click on **9-ExampleDeployWithAdditionalProducts.ps1** and click **Edit**. 
Windows PowerShell ISE will open and we will be able to modify the script before we run it.  
2. Leave, remove, or modify the examples for VisioProRetail and ProjectProRetail.  
3. When you are finished editing the script click **File** and choose **Save**.  
4. Close the Windows PowerShell ISE window.  
5. From the Windows PowerShell console type **.\9-ExampleDeployWithAdditionalProducts.ps1** and press **Enter**.  

###Scenario: Upgrade previous versions of Office to Office ProPlus 
This script will capture existing Office installations and generate a configuration.xml file. 
Previous versions of Office and Office products will be removed using the offscrub scripts. 
Once all products are removed Office ProPlus will be installed using the configuration.xml.  
1. From the PowerShell console type **.\10-ExampleRemovePreviousAndUpgrade.ps1** and press **Enter**.  

##Managed deployments
###Managed deployments using System Center Configuration Manager 
1. From the Configuration Manager server, open Windows PowerShell with elevated privileges.   
2. Change the directory to the Setup-CMOfficeDeployment folder. For example, type **cd C:\PowerShellScripts\Setup-CMOfficeDeployment** and press **Enter**.  
3. Dot source the Setup-CMOfficeDeployment.ps1 script. For example, type **. .\Setup-CMOfficeDeployment.ps1** and press **Enter**. 
Dot sourcing the script allows us to run functions outside of the scope of the script.  
4. Download a channel that will be deployed. .  
	a. Type **Download-CMOfficeChannelFiles –Channels Current –OfficeFilesPath E:\OfficeChannelFiles –Bitness v32**  
5. Create the package.  
	a. Type **Create-CMOfficePackage –Channels Current –Bitness v32 –OfficeSourceFilesPath E:\OfficeChannelFiles -MoveSourceFiles $true**  
	b. If a package called Office 365 ProPlus already exists the script will not run.   
6. Create the deployment program.  
	a. Type **Create-CMOfficeDeploymentProgram –Channels Current –Bitness v32 –DeploymentType DeployWithScript**   
	b. For the DeploymentType you can choose between **DeployWithConfigurationFile** or **DeployWithScript**. 
	Deploying with a configuration file will standardize the deployment for all clients in the collection. 
	Deploying with script will still deploy the channel and bit, but it will preserve the languages and other configurations that may impact the Office installation.   
7. Distribute the package to the distribution points.  
	a. Type **Distribute-CMOfficePackage –DistributionPoint cm.contoso.com**    
	b. Wait for the distribution to finish before creating more programs or deploying the program.  
8. Deploy the program to a collection.  
	a. Type **Deploy-CMOfficeProgram –Collection Accounting –ProgramType DeployWithScript –Channel Current –Bitness v32 –DeploymentPurpose Available**    
9. If you plan to deploy different channels repeat steps 4, 6, 8, and 9. You will need to update the package before step 8. 
For example, if we wanted to create a deployment program for the Deferred channel we would perform the following steps:  
	a. **Download-CMOfficeChannelFiles –Channels Deferred –OfficeFilesPath E:\OfficeChannelFiles –Bitness v32**    
	b. **Create-CMOfficeDeploymentProgram –Channels Deferred –Bitness v32 –DeploymentType DeployWithScript**    
	c. **Update-CMOfficePackage –Channels Deferred –OfficeSourceFilesPath E:\OfficeChannelFiles -MoveSourceFiles $true**. Wait for the distribution to finish before deploying the program to the collection.   
	d. **Deploy-CMOfficeProgram –Collection HR –ProgramType DeployWithScript –Channel Deferred –Bitness v32 –DeploymentPurpose Available**    

###Managed deployments using Group Policy 
1. Copy the Configure-GPOOfficeInstallation folder locally.  
2. Open PowerShell with elevated privileges.  
3. Change the directory to Configure-GPOOfficeInstallation folder. For example, type cd C:\PowerShellScripts\Configure-GPOOfficeInstallation and press Enter.  
4. Dot source the Configure-GPOOfficeInstallation.ps1 script. Type . .\Configure-GPOOfficeInstallation.ps1 and press Enter.
Dot sourcing the script allows us to run functions outside of the scope of the script.  
5. Download the required Office installation files and modify the Office Deployment Tool (ODT) xml files that will be used to deploy Office.
For example, type Download-GPOOfficeInstallation –UncPath \\server1\OfficeChannelFiles -OfficeVersion Office2016 –Bitness 32  
6. Configure an existing group policy. Type Configure-GPOOfficeInstallation –UncPath \\server1\OfficeChannelFiles -GpoName OfficeProPlusDeployments –OfficeVersion Office2016  
7. Refresh the group policy on a client that has the configured GPO and Office ProPlus will install after the next restart.  

##Package an MSI or Executable file to deploy Office ProPlus
The Install Toolkit is an application that will package an Office 365 ProPlus installation into a single Executable or Windows Installer Package (MSI) file. 
The XML configuration file is embedded in the file which allows you to easily distribute Office 365 ProPlus with a custom configuration.  
1. Go to the Configuration XML Editor website located [here](http://officedev.github.io/Office-IT-Pro-Deployment-Scripts/XmlEditor.html).  
2. In the left panel under Tools click on **Install Toolkit**.  
3. Click **Launch Installation**.  
4. Run the **OfficeProPlusInstallGenerator.application** file.  
5. Click **Install** on the Application Install – Security Warning prompt. Accept the security warnings.  
6. The Install Toolkit will open automatically. To begin the process of packaging a deployment file make sure Create new Office 365 ProPlus installer is selected and click Start.  
7. Select the main Office product, Office 365 ProPlus or Office 365 for Business.  
8. Select the edition of Office to install, 32-Bit or 64-Bit.  
9. Click the dropdown under Channel and choose the channel you would like to deploy.  
10. Check the box next to any of the additional Office products you need.  
11. Click **Next**.  
12. Click Add Language and choose any additional languages you may need to install and click OK. If a language other than English needs to be the primary highlight the necessary language and click Set Primary.  
13. Click **Next**.  
14. Use the default Version to install the latest version.  
15. Add a Remote Logging Path, Source Path, or Download Path and click Next.  
16. Deselect any applications that need to be excluded from the deployment and click Next.  
17. Click **Next** in the Optional window.  
18. Updates are enabled by default. Deselect Enable to turn Updates off and continue to step 21.  
19. Verify the Channel is the same as the channel select in step 9.  
20. You may add the Update Path, Target Version and Deadline or leave them as default.  
21. Click **Next**.  
22. Select or deselect the list of additional available options and click **Next**.  
23. Choose MSI or Executable.  
24. If you need to sign the installer using a certificate check the box next to Sign installer and click **Select Certificate** or **Generate Certificate**.  
25. Add a version or leave as default.  
26. Choose Silent install to run the file silently.  
27. Enter the file path to save the generated file or leave as default.  
28. Check Embed Office installation file to include the Office files with the MSI or Exe. Leave this check box blank if you have chosen to use a local Source Path or to install from the Microsoft content delivery network (CDN).  
29. Click Generate and click OK when the process has finished.  


