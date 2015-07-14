#**Duplicate GPO settings from Office 2013 to Office 2016**

In this scenario, we want to automate the process of moving from Office 2013 to Office 2016 while retaining the current set of group policies.  The script will migrate the Office 15 group policies and convert them to Office 16 group policies. 

###**Pre-requisites**

Before running the script, you will need at least the following configuration:

1. Provision a domain controller to manage the Group Policies.
2. Import the Office 2013 Administrative Templates. Information on this process can be found [here](https://www.microsoft.com/en-us/download/details.aspx?id=35554)
3. Provision a client with Office 2013. 
4. Push some group policies to the client.
5. Update the group policy on the client.
6. Upgrade the client to Office 2016.

#####**Update the admx files**

To simulate the Office 2016 admx files make copies of the 2015 admx files for Office (access15.admx, excel15.admx, word15.admx, etc) and change the 15 with 16. Open the admx files in a valid program such as NotePad, and replace the target prefix, namespace, filename, and the reg key strings with 16. 

#####**Update the admx filename**

1. Go to %windir%\PolicyDefinitions.
2. Make a copy of the 15 version of the app and paste it in the same directory.
3. Rename the copied file to *appname*16.admx.

#####**Update the admx file contents**

1. From %windir%\PolicyDefinitions open the new admx files described above in NotePad.
2. Press Ctrl + H or go to Edit > Replace.
3. Next to "Find what" enter *appname*15.
4. Next to " Replace with" enter *appname*16 and click Replace All.

#####**Update the registry data**

1. With the new admx file still open press Ctrl + H or go to Edit > Replace.
2. Next to Find What enter "15.0".
3. Next to Replace What enter "16.0" and click Replace All.

###**Test the script**

1. Import the PowerShell module.

		Import-Module ServerManager
		Add-WindowsFeature GPMC
	
2. Run the script.

		PS C:\> C:\Users\Labadmin\Downloads\Copy-OfficePolicies.ps1
		cmdlet Copy-OfficePolicies.ps1 at command pipeline position 1
		Supply values for the following parameters:
		SourceGPOName: 
	
3. Enter a GPO name.
	
		SourceGPOName: Default Domain Policy

4. Validate the .pol files and Administrative Templates. 
	
		○ The .pol files are located at %systemroot%\sysvol\sysvol\*domain*\Policies\*GUID*\*User or Machine*\.

		○ Open Group Policy Management Editor. The Administrative Templates will have a 2013 copy and a 2016 copy. For example:

	
