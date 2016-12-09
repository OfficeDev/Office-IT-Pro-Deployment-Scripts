#**Install or Upgrade Office 365 ProPlus**

This guide provides the different options for installing or upgrading Office 365 ProPlus. We've divided the guide into 3 sections:  

**Unmanaged deployments** - For organizations that are not using managed deployment software like System Center Configuration Manager and Microsoft Intune. 
Several examples are shown on how to deploy Office ProPlus using PowerShell.  

**Managed deployments** - Deploy Office 365 ProPlus through System Center Configuration Manager using PowerShell commands. In addition, instructions for deploying
Office ProPlus through Group Policy are also outlined.

**Package Office deployments into an MSI or EXE** - Package an Office 365 ProPlus installation into an MSI or Executable file using the Microsoft Office ProPlus Install Toolkit.  

##**Upgrade Office 2013 using PowerShell**

###**Before you run a script**
1. Copy the folder containing the necessary PowerShell scripts and files locally.  
2. Run PowerShell with elevated privileges.  
a. In Windows 10, from the Cortana search box type PowerShell, right click on Windows PowerShell and choose Run as Administrator.
	If you need to enter the credentials of a privileged account, right click Windows PowerShell and choose Pin to Taskbar.
	Hold the Shift key and right click on the Windows PowerShell icon and choose Run as different user.  
b. In Windows 8.1, go to the Start Menu and type powershell. Right click on Windows PowerShell and choose Run as Administrator.
	If you need to enter the credentials of a privileged account, right click Windows PowerShell and choose Pin to Taskbar, and open 	 the Desktop.
	Hold the Shift key and right click on the Windows PowerShell icon, and choose Run as different user.  
c. In Windows 7, click the Start button and type powershell. Right click on Windows PowerShell and choose Run as administrator.
	If you need to enter the credentials of a privileged account hold the Shift key and right click on Windows PowerShell and choose         Run as different user.  
3. Change the directory to the Deploy-OfficeClickToRun folder.  
	a. For example, type **cd C:\Deploy-OfficeClickToRun and press Enter.**  
4. Set the execution policy to unrestricted or bypass.  
	a. In the PowerShell console, type **Set-ExecutionPolicy Unrestricted**, press **Enter**, and accept the policy change request. 
	For more information about execution policies please visit http://go.microsoft.com/fwlink/?LinkID=135170.  
Note: By typing **.\** before a script name lets PowerShell know to execute a script located in the current directory. 
PowerShell also uses IntelliSense which will allow you to start typing the name of a script, cmdlet, or function and press tab to finish the name. 
Use PowerShell to install or upgrade to Office ProPlus

###Manually upgrade Office 2013 to Office 365 ProPlus
This script is meant to be a template for upgrading existing Office installations. It will capture the existing Office 2013 installation and generate a configuration.xml file. 
Office 2013 products will be removed using the offscrub scripts and when all products are removed Office 365 ProPlus will be installed using the configuration.xml.  

1. Right click on **.\10-ExampleRemovePreviousAndUpgrade.ps1** and click Edit. PowerShell ISE will open that will allow us to make changes to the script.  
2. Make any necessary changes to fit your deployment. For example, if you are deploying the Current channel replace the vairable **-Channel Deferred** with **-Channel Current**.  
3. After you've modified the script to fit your Office installation requirements **save** and **close** the script.  
4. From the PowerShell console type **.\10-ExampleRemovePreviousAndUpgrade.ps1** and press **Enter**.  

##Upgrade Office 2013 using a managed deployment

###Upgrade Office 2013 to Office 365 ProPlus using System Center Configuration Manager
1. Before we begin, determine if you need to modify the script **CM-ExampleRemovePreviousAndUpgrade.ps1** located inside of the DeploymentFiles folder. Right click on the script and choose **Edit**.  
2. The script is designed to be as dynamic as possible by capturing existing Office installations on the computer and generating the configuration.xml.  
Make any necessary changes that align with your deployment criteria, then save and close the file. 
3. From the Configuration Manager server, open Windows PowerShell with elevated privileges.   
4. Change the directory to the Setup-CMOfficeDeployment folder. For example, type **cd C:\PowerShellScripts\Setup-CMOfficeDeployment** and press **Enter**.  
5. Dot source the Setup-CMOfficeDeployment.ps1 script. For example, type **. .\Setup-CMOfficeDeployment.ps1** and press **Enter**. 
Dot sourcing the script allows us to run functions outside of the scope of the script.  
6. Download a channel that will be deployed. .  
	a. Type **Download-CMOfficeChannelFiles –Channels Current –OfficeFilesPath E:\OfficeChannelFiles –Bitness v32**  
7. Create the package.  
	a. Type **Create-CMOfficePackage –Channels Current –Bitness v32 –OfficeSourceFilesPath E:\OfficeChannelFiles -MoveSourceFiles $true**  
	b. If a package called Office 365 ProPlus already exists the script will not run.   
8. Create the deployment program.  
	a. Type **Create-CMOfficeDeploymentProgram –Channels Current –Bitness v32 –DeploymentType DeployWithScript**   
	b. For the DeploymentType you can choose between **DeployWithConfigurationFile** or **DeployWithScript**. 
	Deploying with a configuration file will standardize the deployment for all clients in the collection. 
	Deploying with script will still deploy the channel and bit, but it will preserve the languages and other configurations that may impact the Office installation.   
9. Distribute the package to the distribution points.  
	a. Type **Distribute-CMOfficePackage –DistributionPoint cm.contoso.com**    
	b. Wait for the distribution to finish before creating more programs or deploying the program.  
10. Deploy the program to a collection.  
	a. Type **Deploy-CMOfficeProgram –Collection Accounting –ProgramType DeployWithScript –Channel Current –Bitness v32 –DeploymentPurpose Available**    
11. If you plan to deploy different channels repeat steps 6, 8, 10, and 11. You will need to update the package before step 8. 
For example, if we wanted to create a deployment program for the Deferred channel we would perform the following steps:  
	a. **Download-CMOfficeChannelFiles –Channels Deferred –OfficeFilesPath E:\OfficeChannelFiles –Bitness v32**    
	b. **Create-CMOfficeDeploymentProgram –Channels Deferred –Bitness v32 –DeploymentType DeployWithScript -ScriptName CM-ExampleRemovePreviousAndUpgrade.ps1**    
	c. **Update-CMOfficePackage –Channels Deferred –OfficeSourceFilesPath E:\OfficeChannelFiles -MoveSourceFiles $true**. Wait for the distribution to finish before deploying the program to the collection.   
	d. **Deploy-CMOfficeProgram –Collection HR –ProgramType DeployWithScript –Channel Deferred –Bitness v32 –DeploymentPurpose Available**    

###Upgrade Office 2013 using using Group Policy 
1. Copy the Configure-GPOOfficeInstallation folder locally.  
2. Open PowerShell with elevated privileges.  
3. Change the directory to Configure-GPOOfficeInstallation folder. For example, type **cd C:\PowerShellScripts\Configure-GPOOfficeInstallation** and press **Enter**.  
4. Dot source the Configure-GPOOfficeInstallation.ps1 script. Type **. .\Configure-GPOOfficeInstallation.ps1** and press **Enter**.
Dot sourcing the script allows us to run functions outside of the scope of the script.  
5. Download the required Office installation files and modify the Office Deployment Tool (ODT) xml files that will be used to deploy Office.
For example, type **Download-GPOOfficeInstallation –UncPath \\server1\OfficeChannelFiles -OfficeVersion Office2016 –Bitness 32**  
6. Configure an existing group policy. Type **Configure-GPOOfficeInstallation –UncPath \\server1\OfficeChannelFiles -GpoName OfficeProPlusDeployments –OfficeVersion Office2016**    
7. Refresh the group policy on a client that has the configured GPO and Office ProPlus will install after the next restart.  