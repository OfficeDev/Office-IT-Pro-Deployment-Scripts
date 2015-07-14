#**Check Remote Office Versions**

This PowerShell Function will query the local or remote workstations to find the version of Office that is installed.   

###**Pre-requisites**

1. Remote Windows Management Instrumentation (WMI) connectivity and permissions to any remote computers you are querying

###**Running the script**

1. Open a PowerShell console.

		From the Run dialog type PowerShell 
		
2. Change directory to the location where the PowerShell Script is saved.

		Example: cd C:\PowerShellScripts
		
2. Run the Script. With no parameters specified the script will return the locally installed Office Version.

		Type .\Get-OfficeVersion.ps1
		Press Enter and then if Microsoft Office is installed locally it should display
	
3. Run the Script against a remote computer. 

		Run the PowerShell script Get-OfficeVersion.ps1 -ComputerName Client01
	

	

