#**Remove previous versions of Office**

This PowerShell Script will create remove the local MSI installations of Office 2013 and older. The Offscrub vbs scripts are used to remove the MSI installations of Office products.

###**Using the Offscrub scripts**

The Offscrub vbs scripts can be used to automate the removal of Office products. The scripts will uninstall the existing Office products regardless of the current health of the installation. The Remove-PreviousOfficeInstalls.ps1 script will determine which version of Office is currently installed and will call the appropriate Offscrub vbs script to remove the Office products installations.

The Offscrub vbs files included are:

* **OffScrub03.vbs** - Used to remove Office 2003 products.
* **OffScrub07.vbs** - Used to remove Office 2007 products.
* **OffScrub10.vbs** - Used to remove Office 2010 products.
* **OffScrub_O15msi.vbs** - Used to remove Office 2013 MSI products.

More information can be found at: https://blogs.technet.microsoft.com/odsupport/2011/04/08/how-to-obtain-and-use-offscrub-to-automate-the-uninstallation-of-office-products/

###Example

1. Open an elevated PowerShell console.

		From the Run dialog type PowerShell 
		
2. Change directory to the location where the PowerShell Script is saved.

		Example: cd C:\PowerShellScripts
		
3. Dot-Source the Remove-PreviousOfficeInstalls function into your current session.

		Type . .\Remove-PreviousOfficeInstalls.ps1
		
		By including the additional period before the relative script path you are 'Dot-Sourcing' 
		the PowerShell function in the script into your PowerShell session which will allow you to 
		run the function from the console.

3. Run the function.

		Type  **Remove-PreviousOfficeInstalls**
		
		The version of Office will be detected automatically and the appropriate Offscrub file will be used to remove any Office products. If Office is not detected on the client the script will notify the admin and stop running.
			

	

	

