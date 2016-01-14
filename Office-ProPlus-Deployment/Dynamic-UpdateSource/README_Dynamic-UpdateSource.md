#Remove Office Click-to-Run

This PowerShell function will update a configuration XML file to add or update the "SourcePath" attribute

###Example

1. Open a PowerShell console.

		From the Run dialog type PowerShell 

2. Change directory to the location where the PowerShell Script is saved.

		Example: cd C:\PowerShellScripts

3. Dot-Source the Remove-OfficeClickToRun function into your current session.

		Type . .\Dynamic-UpdateSource.ps1
		By including the additional period before the relative script path you are 'Dot-Sourcing' 
		the PowerShell function in the script into your PowerShell session which will allow you to 
		run the function from the console.
		
4. Run the function against the local computer, be sure to include the parameters TargetFilePath and UpdateSourcePath.

		Dynamic-UpdateSource -TargetFilePath "\\server\msoffice\configuration.xml" -UpdateSourcePath
		"\\server\msoffice\site.csv"

5. Have a csv file which have the header row "ADSite,source" note the sample content below, this file can either have a .csv or .txt extension

		ADSite,source
		MS-HQ1,\\MS-HQ1\MSOfficeSource
		MS-HQ2,\\Server1\MSOfficeSource
		MS-HQ3,\\Server2\MSOfficeSource
		
