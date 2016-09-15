### Configure GPO Office Installation
This script will configure an existing Active Directory Group Policy to silently install Office 2013 Click-To-Run on computer startup.

**IT Pro Scenario:** For organizations that aren't using SCCM or another managed deployment system, this script will install Office using a Group Policy. 

[README](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/wiki/README_New-GPOOfficeInstallation)

1. Open an elevated PowerShell console.

		From the Run dialog type PowerShell and run as administrator

2. Change directory to the location where the PowerShell Script is saved on your local machine.

		Example: cd C:\PowerShellScripts

3. Dot-Source the Configure-GPOOfficeInstallation function into your current session.

		Type . .\Configure-GPOOfficeInstallation.ps1
		
		By including the additional period before the relative script path you are 'Dot-Sourcing' 
		the PowerShell function in the script into your PowerShell session which will allow you to 
		run the function from the console.

4. Run the Configure-GPOOfficeInstallation functions to download and then configure the Group Policy

		Example: Download-GPOOfficeInstallation -Bitness v32

			 Configure-GPOOfficeInstallation -GPOName %seattleusers% -UncPath %\\contosofs\office\installers% -ConfigFileName %configuration.xml%






