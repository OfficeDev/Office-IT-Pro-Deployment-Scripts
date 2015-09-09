#**Nuke-Office**

This PowerShell Script will create a scheduled task on a remote PC to completely remove whatever version of Office is present, including in the registry.   
The Script will also notify you of any Add-Ins present on the PC.

###**Pre-requisites**

1. Remote Windows Management Instrumentation (WMI) connectivity and permissions to any remote computers you are querying. 

2. Be sure that your PowerShell execution policy allows running scripts.
		
		Set-ExecutionPolicy Unrestricted

###**Examples**

1. Open a PowerShell console.

		From the Run dialog type PowerShell 
		
2. Change directory to the location where the PowerShell Script is saved.

		Example: cd C:\PowerShellScripts
		
2. Run the Script. With no parameters specified the script will remove the locally installed Office Version.

		Type . .\Nuke-Office.ps1
		Press Enter and then if Microsoft Office is installed locally it should display. 
		By including the additional period before the relative script path you are 'Dot-Sourcing' 
		the PowerShell function in the script into your PowerShell session which will allow you to 
		run the function from the console.
	
3. Run the Script against a remote computer. 

		Type Nuke-Office -ComputerName Client01

4. Run the Script against multiple remote computers. 

		Type Nuke-Office -ComputerName Client01,Client02
	

	

