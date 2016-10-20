### Generate ODT Configuration XML
This PowerShell Function generates the Office Deployment Tool Configuration XML based on the current state of the workstation and the parameters specified for the Function.

[README](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/wiki/README_Generate-OfficeDeploymentTool(ODT)ConfigurationXML)


Open a PowerShell console.

		From the Run dialog type PowerShell 

2. Change directory to the location where the PowerShell Script is saved.

		Example: cd C:\PowerShellScripts

3. Dot-Source the Generate-ODTConfigurationXML into your current session.

		Type . .\Generate-ODTConfigurationXML
		
		By including the additional period before the relative script path you are 'Dot-Sourcing' 
		the PowerShell function in the script into your PowerShell session which will allow you to 
		run the function from the console.

4. Run the Generate-ODTConfigurationXML function to create the configurationXML to be used with the odt installer.

		Example: Generate-ODTConfigurationXML -Languages %CurrentOfficeLanguages,OSLanguages,OSandUserLanguages,AllInUseLanguages% -TargetFilePath %c:\O365\installer\%

