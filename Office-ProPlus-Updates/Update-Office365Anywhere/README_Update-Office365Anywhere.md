##**Update Office 365 Anywhere**

This function is designed to provide a way for Office Click-To-Run clients to have the ability to update themselves from a managed network source or from the Internet depending on the availability of the primary update source.  The idea behind this is if users have laptops and are mobile they may not receive updates if they are not able to be in the office on a regular basis. 

How it works is if the configured Update source for the Office Click-To-Run is configured to a network resource like a network share or network web source then it will first check to see if the source is available.  If the source is not available the script will assume the client is not on the corporate network and it will update the client from the Internet using the Microsoft Office Content Delivery Network (CDN). This ensures all the mobile clients do not have to download updates from the Internet while they are on the corporate network while ensuring they are able to still receive the updates if they are not regularly in the office.

This functionality is available with this function but it's use can be controlled by the parameter -EnableUpdateAnywhere.  This function also provides a way to initiate an update and the script will wait for the update to complete before exiting. Natively starting an update executable does not wait for the process to complete before exiting and in certain scenarios it may be useful to have the update process wait for the update to complete.

The script considers the primary Update source whatever is configured in the following registry values.

		Office 2013 - HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration\UpdateUrl
		Office 2016 - HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Configuration\UpdateUrl

The configuration of this attribute is not in scope of this script but the there are ways to manage this update source to include a script for SCCM called [Setup-SCCMOfficeUpdates](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/tree/Development/Office-ProPlus-Updates/Setup-SCCMOfficeUpdates)

###**Running the script**

1. Open an Elevated PowerShell Console(see, [Starting Windows PowerShell](https://technet.microsoft.com/en-us/library/hh857343.aspx)):

		From the Run dialog type PowerShell 
		
2. Change directory to the location where the PowerShell Script is saved.

		Example: cd C:\PowerShellScripts
		

4. Type the following in the elevated PowerShell Session

		 . .\Update-Office365Anywhere.ps1
         

