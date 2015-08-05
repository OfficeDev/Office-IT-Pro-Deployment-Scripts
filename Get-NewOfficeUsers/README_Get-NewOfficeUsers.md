#Get New Office Users

Two scripts to identify licensed Office 365 users and track the dates they were enabled or disabled.

###Pre-requisites

The script requires the Azure Active Directory Services Module, previously known as the Microsoft Online Services Module.

The module information can be found at https://technet.microsoft.com/en-us/library/jj151815.aspx.

###Update-UserLicenseData

Finds all of the MSOLUsers that are licensed with the specified plan and stores them in a CSV.
An extra field is populated in the CSV that specifies the earliest point at which the CSV
knew the user was licensed (LicensedAsOf field). If a user doesn't show up in the licensed list at a later date
the function takes note and populates another field in the CSV with that date (DelicensedAsOf).

####Examples

1. Open a PowerShell console.

		From the Run dialog type PowerShell.
	
2. Change directory to the location where the PowerShell Script is saved.

		Example: cd C:\PowerShellScripts
	
3. Run the Update-UserLicenseData.ps1 script.

		Type . .\Update-UserLicenseData.ps1 -ServiceName "OFFICESUBSCRIPTION" -CSVPath "C:\Licensed Users\LicensedUsers.csv"

		By including the additional period before the relative script path you are 'Dot-Sourcing' 
		the PowerShell function in the script into your PowerShell session which will allow you to 
		run the function from the console.
		
4. Run the Update-UserLicenseData.ps1 script and specify a username and password.

		Type . .\Update-UserLicenseData.ps1 -ServiceName "OFFICESUBSCRIPTION" -CSVPath "C:\Licensed Users\LicensedUsers.csv" -Username "administrator" -Password "adminpassword"
		
###Get-RecentlyLicensedUsers

Get a list of users the were licensed after the specified date according to the specified CSV.
It is important to have run the Update-UserLicenseData.ps1 prior to using this script.

####Examples

1. Open a PowerShell console.

		From the Run dialog type PowerShell.
	
2. Change directory to the location where the PowerShell Script is saved.

		Example: cd C:\PowerShellScripts
	
3. Run the Get-NewOfficeUsers.ps1 script.

		Type . .\Get-RecentlyLicensedUsers.ps1 -CutOffDate (Get-Date "2015-8-5") -CSVPath "C:\Licensed Users\LicensedUsers.csv"
