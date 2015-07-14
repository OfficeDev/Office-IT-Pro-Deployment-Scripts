#**Duplicate GPO settings from Office 2013 to Office 2016**

In this scenario, we want to automate the process of moving from an existing version of Office to a newer version while retaining the current set of group policies. The script will pull the policy file paths of the GPO name, GUID, domain, user and machine registry.pol files, and the PolicyDefinitions. It pulls a list of the Office target version admx files and the definitions of the admx files. It will then copy the policy information if the keypath exists in the target version which is checked against the definitions pulled from the admx files.  

###**Pre-requisites**

Before running the script, you will need at least the following configuration:

1. A domain controller managing the Group Policies.
2. SourceVersion Administrative templates installed.
3. Group Policies from the SourceTarget version deployed.
3. TargetVersion Administrative templates installed.

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
	
		SourceGPOName: "Office Group Policy"

4. Validate the .pol files and Administrative Templates. 
	
		○ The .pol files are located at %windir%\sysvol\sysvol\domain\Policies\GUID\User or Machine\.

		○ Open Group Policy Management Editor. The Administrative Templates will have a SourceVersion copy and a TargetVersion copy. Verify the settings from the SourceVersion are now set in the TargetVersion.

	
