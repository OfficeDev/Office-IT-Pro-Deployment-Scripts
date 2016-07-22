#Remove Office Click-to-Run

This PowerShell function will create a configuration xml file and uninstall Office Click-to-Run via the Office Deployment Tool (ODT).

###Example

1. Open a PowerShell console.

		From the Run dialog type PowerShell 

2. Change directory to the location where the PowerShell Script is saved.

		Example: cd C:\PowerShellScripts

3. Dot-Source the Remove-OfficeClickToRun function into your current session.

		Type . .\Remove-OfficeClickToRun.ps1
		By including the additional period before the relative script path you are 'Dot-Sourcing' 
		the PowerShell function in the script into your PowerShell session which will allow you to 
		run the function from the console.
		
4. Run the function against the local computer.

		Remove-OfficeClickToRun
