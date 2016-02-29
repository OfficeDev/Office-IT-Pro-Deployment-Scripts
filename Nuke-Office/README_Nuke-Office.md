#**Nuke-Office**

This PowerShell Script will create a scheduled task on a remote PC to completely remove any version of Office present.   
The Script will also notify you of any Office Add-Ins present on each of the PCs.  Administrator permission is required 
on the PCs where the tasks will be run.  

###**Pre-requisites**

1. Remote Windows Management Instrumentation (WMI) connectivity and Admin permissions to any remote computers you are querying. 

2. Make sure that your local PowerShell execution policy allows running scripts.
		
		Set-ExecutionPolicy Unrestricted

###**Instructions and Examples**

1. Open a PowerShell console.

		From the Run dialog type PowerShell 
		
2. Change directory to the location where the PowerShell Script is saved.

		Example: cd C:\PowerShellScripts
		
2. Run the Script. With no parameters specified the script will remove the local installed version of Office from the PC.

		Type  .\Nuke-Office.ps1
			
3. Run the Script against a remote computer. 

		Type Nuke-Office -ComputerName Client01

4. Run the Script against multiple remote computers. 

		Type Nuke-Office -ComputerName Client01,Client02

5. Run the script against a remote computer and supply Admin credentials.

		Type Nuke-Office -ComputerName Client01 -credential
	

	

