### Check Office Pro Plus Ports

Checks the availability of the various remote resources needed to install Office 365

###**Running the script**

1. Open a PowerShell console:

		From the Run dialog type PowerShell.
	
2. Change directory to the location where the PowerShell Script is saved.

		Example: cd C:\PowerShellScripts
	
3. Run the Check-OfficeProPlusPorts.ps1 script.

		Type . .\Check-OfficeProPlusPorts

		By including the additional period before the relative script path you are 'Dot-Sourcing' 
		the PowerShell function in the script into your PowerShell session which will allow you to 
		run the function from the console.

4. Verify that there are no "failed" port checks. 

5. If there are ports blocked, unblock them and rerun this script to verify you pass the port requirements.  


**IT Pro Scenario:** Organizations that are upgrading Office to 2016 and would like to ensure the machines firewall settings are correct before installation.