#**Nuke-Office**

This PowerShell Script will create remote the local MSI installations of Office 2013 and older. The script using the Offscrub script in order to remove the installations of Office.

###**Pre-requisites**

1. Remote Windows Management Instrumentation (WMI) connectivity and Admin permissions to any remote computers you are querying. 

2. Make sure that your local PowerShell execution policy allows running scripts.
		
		Set-ExecutionPolicy Unrestricted

###**Instructions and Examples**

1. Open a PowerShell console.

		From the Run dialog type PowerShell 
		
2. Change directory to the location where the PowerShell Script is saved.

		Example: cd C:\PowerShellScripts
		
2. Run the Script. Script will remove MSI installations of Office 2013 and older

		Type  .\Remove-PreviousOfficeInstalls.ps1
			

	

	

