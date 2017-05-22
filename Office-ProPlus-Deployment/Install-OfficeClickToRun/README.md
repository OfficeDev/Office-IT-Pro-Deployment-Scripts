# **Install Office Click-To-Run**

This PowerShell function will take an existing Office Click-To-Run configuration xml and deploy it using the Office Deployment Tool (ODT). 

### **Examples**

1. Open a PowerShell console.

		From the Run dialog type PowerShell 

2. Change directory to the location where the PowerShell Script is saved.

		Example: cd C:\PowerShellScripts

3. Dot-Source the Generate-ODTConfigurationXML function into your current session.

		Type . .\Install-OfficeClickToRun.ps1
		By including the additional period before the relative script path you are 'Dot-Sourcing' 
		the PowerShell function in the script into your PowerShell session which will allow you to 
		run the function from the console.

4. Run the function against the local computer

		Install-OfficeClickToRun -TargetFilePath configuration.xml 

6. In order to create a complete solution there are two other scripts in this GitHub Repository.  

	[Generate-ODTConfigurationXML](../Generate-ODTConfigurationXML) - This script provides a function to Generate the Office Click-To-Run configuration xml file.
	
	[Edit-OfficeConfigurationFile](../Edit-OfficeConfigurationFile) - This script provides functions to edit the Office Click-To-Run configuration file.

[![Analytics](https://ga-beacon.appspot.com/UA-70271323-4/README_Install-OfficeClickToRun?pixel)](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts)
