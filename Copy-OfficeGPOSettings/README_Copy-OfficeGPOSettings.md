#**Copy GPO settings from Office 2013 to Office 2016**

Automates the process of copying configured Office 2013 Group Policy Settings to the Office 2016 Group Policy Settings. 

###**Pre-requisites**

Before running the script, you will need at least the following requirements

1. The script must be run from a Domain Controller in the domain of the Group Policy you with to modify.
2. The Administrative templates of the source Office version installed on the Domain Controller.
2. The Administrative templates of the target Office version installed on the Domain Controller.

###**Running the script**

1. Open a PowerShell console:

		From the Run dialog type PowerShell.
	
2. Change directory to the location where the PowerShell Script is saved.

		Example: cd C:\PowerShellScripts
	
3. Run the Copy-OfficeGPOSettings.ps1 script.

		Type . .\Copy-OfficeGPOSettings.ps1 -SourceGPOName "Office Group Policy Name"

		By including the additional period before the relative script path you are 'Dot-Sourcing' 
		the PowerShell function in the script into your PowerShell session which will allow you to 
		run the function from the console.

4. Verify that the all of the Office 2013 settings have been copied to the Office 2016 settings 


	
