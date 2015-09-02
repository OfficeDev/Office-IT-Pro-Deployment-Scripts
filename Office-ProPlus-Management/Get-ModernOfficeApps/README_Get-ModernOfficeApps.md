#Check for Modern Office Apps

List the Modern Office apps installed on the local or remote computers

###**Pre-requisites**

1. Remote Windows Management Instrumentation (WMI) connectivity and permissions to any remote computers you are querying. 

###Examples

####Check for modern apps installed on remote computers:

1. Open a PowerShell console.

            From the Run dialog type PowerShell
            
2. Change the directory to the location where the PowerShell Script is saved.

            Example: cd C:\PowerShellScripts
            
3. Run the Script. With no parameters specified the script will return the locally installed Office Version.

           Type . .\Get-ModernOfficeApps.ps1

           By including the additional period before the relative script path you are 'Dot-Sourcing' 
           the PowerShell function in the script into your PowerShell session which will allow you to 
           run the function 'Get-ModernOfficeApps' from the console.
	
4. Run the script for specified computers or against an array you have created.

            Example: .\Get-ModernOfficeApps.ps1 -ComputerName Client1,Client2
            


