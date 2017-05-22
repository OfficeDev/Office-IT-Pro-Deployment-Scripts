# Dynamically Set Source Path, Office Click-to-Run

This PowerShell function works with the other deployment scripts to provide a way to dynamically set the SourcePath for the Office Click-To-Run configuration xml.  The script uses the computers Active Directory site to lookup the SourcePath location from the LookupFilePath.  The LookupFilePath is a CSV file that you will need to populate with the Active Directory site and corresponding SourcePath.

### Example

1. Open a PowerShell console.

		From the Run dialog type PowerShell 

2. Change directory to the location where the PowerShell Script is saved.

		Example: cd C:\PowerShellScripts

3. Dot-Source the Dynamic-UpdateSource function into your current session.

		Type . .\Dynamic-UpdateSource.ps1
		By including the additional period before the relative script path you are 'Dot-Sourcing' 
		the PowerShell function in the script into your PowerShell session which will allow you to 
		run the function from the console.
		
5. Update the SourcePathLookup.csv with the Active Directory Site name and the update source for that site.
		ADSite,source
		Site1,\\Site1Server\OfficeSource
		Site2,\\Site2Server\OfficeSource
		Site3,\\Site3Server\OfficeSource
		
4. Run the function against the local computer, be sure to include the parameters TargetFilePath and IncludeUpdatePath.

		Dynamic-UpdateSource -TargetFilePath "\\server\msoffice\configuration.xml" -IncludeUpdatePath $true


[![Analytics](https://ga-beacon.appspot.com/UA-70271323-4/README_Dynamic-UpdateSource?pixel)](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts)
