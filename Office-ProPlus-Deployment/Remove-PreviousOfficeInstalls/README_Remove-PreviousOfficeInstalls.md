#**Remove previous versions of Office**

This PowerShell Script will create remote the local MSI installations of Office 2013 and older. The Offscrub vbs scripts are used to remove the MSI installations of Office products.

###**Using the Offscrub scripts**

The Offscrub vbs scripts can be used to automate the removal of Office products. The scripts will uninstall the existing Office products regardless of the current health of the installation. The Remove-PreviousOfficeInstalls.ps1 script will determine which version of Office is currently installed and will call the appropriate Offscrub vbs script to remove the Office products installations.

The Offscrub vbs files included are:

* **OffScrub03.vbs** - Used to remove Office 2003 products.
* **OffScrub07.vbs** - Used to remove Office 2007 products.
* **OffScrub10.vbs** - Used to remove Office 2010 products.
* **OffScrub_O15msi.vbs** - Used to remove Office 2013 MSI products.

###Examples

1. Open a PowerShell console.

		From the Run dialog type PowerShell 
		
2. Change directory to the location where the PowerShell Script is saved.

		Example: cd C:\PowerShellScripts
		
3. Dot-Source the Remove-PreviousOfficeInstalls function into your current session.

		Type . .\Remove-PreviousOfficeInstalls.ps1
		
		By including the additional period before the relative script path you are 'Dot-Sourcing' 
		the PowerShell function in the script into your PowerShell session which will allow you to 
		run the function from the console.

3. Run the Script. Script will remove MSI installations of Office 2013 and older

		Type  Remove-PreviousOfficeInstalls
			

	

	

