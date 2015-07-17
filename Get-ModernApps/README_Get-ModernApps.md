#Check for Modern Apps

Remotely verify the modern apps installed on computers in a domain.

###Examples

#####Check for modern apps installed on all computers:

1. Open a PowerShell console.

            From the Run dialog type PowerShell
            
2. Change the directory to the location where the PowerShell Script is saved.

            Example: cd C:\PowerShellScripts
            
3. Type Get-ModernAppsRemotely.ps1 and press Enter.

4. Enter the credentials of an administrator account.

####Check for modern apps installed on specified computers:

1. Open a PowerShell console.

            From the Run dialog type PowerShell
            
2. Change the directory to the location where the PowerShell Script is saved.

            Example: cd C:\PowerShellScripts
            
3. Run the script for specified computers or against an array you have created.

            Example: .\Get-ModernAppsRemotely.ps1 -ComputerNames ( $myArray )
            
4. Enter the credentials of an administrator account if requested.


