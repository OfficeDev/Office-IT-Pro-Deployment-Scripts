# **Move Office 2016 ShortCuts**

This PowerShell Script will move the Start Menu shortcuts created by Office 365 ProPlus client into a sub folder instead of the default root location

### Example

1. Open a PowerShell console.

		From the Run dialog type PowerShell 
		
2. Change directory to the location where the PowerShell Script is saved.

		Example: cd C:\PowerShellScripts
		
3. Dot-Source the Remove-PreviousOfficeInstalls function into your current session.

		Type . .\Move-Office365ClientShortCuts.ps1
		
		By including the additional period before the relative script path you are 'Dot-Sourcing' 
		the PowerShell function in the script into your PowerShell session which will allow you to 
		run the function from the console.

3. Run the function.

		Type  Move-Office365ClientShortCuts
		
	

	

