# **Configuration Manager Office Add-in Reporting**

Use PowerShell scripts in this GitHub repository to create a hardware inventory in System Center Configuration Manager that will inventory installed Office add-ins. 
A SQL table is created by importing a custom mof file. Several Office Add-in reports are already available to import.

More information on MOF files can be found here: https://technet.microsoft.com/en-us/library/cc180827.aspx

## **Clone or download the required files locally:**
*	CMOfficeAddinReports (contains ready to use reports that can be imported into Configuration Manager)
*	ScriptFiles (contains the scripts required)
*	1-ExampleSetupCMOfficeAddinPackage.ps1
*	2-ExampleSetupCMOfficeAddinPackageWithScheduledTask.ps1
*	Custom_OfficeAddins.mof
*	Setup-CMOfficeAddinPackage

## Add the new hardware inventory class
### Step 1: Import the mof file
1. From a Configuration Manager console, go to **Administration > Client Settings**.
2. Right click the Client Setting that will be used to enable the hardware inventory and select **Properties**.
3. For new client settings or client settings that do not have Hardware Inventory enabled, check the box next to **Hardware Inventory**.
4. In the left pane, click **Hardware Inventory**.
5. Click **Set Classes ...**
6. Click **Import...**
7. Select **Custom_OfficeAddin.mof** and click **Open**.
8. Click **OK** until the Client Settings window closes. 

### Step 2: Deploy the Client Setting
If the client setting from Step 1 has already been deployed to a device collection, skip this section continue to the next section.
1. From **Administration > Client Settings**, right click the Client Setting to deploy and choose **Deploy**.
2. From **Device Collections**, select the Device Collection and click **OK**.

## Create a package, program and deployment
### Step 1: Create a package that will host the required scripts
A PowerShell script must run on a client device that will create a new WMI Class which will contain the necessary Office add-in data for that device.
The process to create a package, program and deployment can be automated by using the Setup-CMOfficeAddinPackage.ps1 script. Example scripts have been created to show the scope of the entire process.

1. Open a PowerShell console with elevated privileges.  
2. Change the directory to the location of the downloaded files. For example,  
	`cd C:\OfficeAddinScripts`
3. Dot source the Setup-CMOfficeAddinPackage script. For example,  
	`. .\Setup-CMOfficeAddinPackage.ps1`
4. Type the cmdlet to create and configure the new package. For example,   
	`Create-CMOfficeAddinPackage -PackageName "Update Office add-in repository" -ScriptFilesPath "C:\OfficeAddinScripts" -MoveScriptFiles $true`

### Step 2: Create a program to deploy the PowerShell scripts to the clients
There are two different functions that can be used to create the programs, Create-CMOfficeAddinProgram and Create-CMOfficeAddinTaskProgram.
* **Create-CMOfficeAddinProgram** is used to create a program that will run the Get-OfficeAddins.ps1 script once.
* **Create-CMOfficeAddinTaskProgram** is used to create a program that will create a scheduled task on the target device. The scheduled task will run the Get-OfficeAddins.ps1 script on the target device once a week.
Use one of the following to create the program:

1. Type the cmdlet to create a basic program. For example,  
	`Create-CMOfficeAddinProgram -PackageName "Update Office add-in repository" -ProgramName "OfficeAddinQuery"`
2. Type the cmdlet to create a scheduled task program. For example,  
	`Create-CMOfficeAddinTaskProgram -PackageName "Update Office add-in repository" -ProgramName "Update with Scheduled Task" -UseRandomStartTime $true -RandomTimeStart "06:00" -RandomTimeEnd "18:00"`

### Step 3: Distribute the package to a distribution point or distribution point group
It's important to let the package finish distributing before moving on the creating the deployment. Use the -WaitForDistributionToFinish switch to show the distribution status.

1. Type the cmdlet to distribute the package. For example,  
	`Distribute-CMOfficeAddinPackage -PackageName "Update Office add-in repository" -DistributionPoint CM01.CONTOSO.COM -WaitForDistributionToFinish $true`

### Step 4: Deploy the program to a device collection
1. Type the cmdlet to create the deployment. For example,  
	`Deploy-CMOfficeAddinProgram -PackageName "Update Office add-in repository" -ProgramName "Update with Scheduled Task" -Collection "All Desktop and Server Clients" -DeploymentPurpose Required`

## Import the reports
Several reports have already been created and customized to show Office add-in data in Configuration Manager. The reports are located [here](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/tree/master/Office-ProPlus-Management/Get-OfficeAddins/CMOfficeAddinReports). 
Use your preferred method to import the reports into the SQL Server Reporting Services.
