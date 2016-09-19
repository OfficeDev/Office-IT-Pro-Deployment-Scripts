### Install Office Click-To-Run
This PowerShell function will install Office Click-To-Run 

[README](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/wiki/README_Install-OfficeClickToRun)


###Links
Group Policy Management Console - https://technet.microsoft.com/en-us/library/cc753298.aspx
WMI filtering - https://technet.microsoft.com/en-us/library/cc779036(v=ws.10).aspx

###**Examples**

1. Open a PowerShell console.

		From the Run dialog type PowerShell 

2. Change directory to the location where the PowerShell Script is saved.

		Example: cd C:\PowerShellScripts

3. Dot-Source the Install-OfficeClickToRun into your current session.

		Type **. .\Install-OfficeClickToRun.ps1**

4. Run the function with the appropriate variables	

		Example:  Install-OfficeClickToRun -targetfilepath %c:\O365\installer% -OfficeVersion Office2016 -WaitForInstallToFinish $true -ConfigurationXML c:\O365\installer\configuration.xml