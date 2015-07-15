#Check for Modern Apps

The purpose of this script is to remotely verify the modern apps installed on computers in a domain.

###Examples

#####Check for modern apps installed on all computers:

1. Open a PowerShell console.
    From the Run dialog type PowerShell
2. Change the directory to the location where the PowerShell Script is saved.
    Example: cd C:\PowerShellScripts
3. Type Get-ModernAppsRemotely.ps1 and press Enter.
3. Enter the credentials of an administrator account.

####Check for modern apps installed on specified computers:

1. Open a PowerShell console.
    From the Run dialog type PowerShell
2. Change the directory to the location where the PowerShell Script is saved.
    Example: cd C:\PowerShellScripts
3. Create a variable for an array of computers
    Example: $myarr = ("Computer1","Computer2")
4. Run the script against the new array
    Example: .\Get-ModernAppsRemotely.ps1 -ComputerNames ( $myarray )
5. Enter the credentials of an administrator account.


