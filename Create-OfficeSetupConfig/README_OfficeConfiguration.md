#Office Configuration XML Builder

Functions for building the configuration xml file for click to run office products
Functions included:
	New-OfficeConfiguration
	Add-Product
	Remove-Product
	Set-Updates
	Set-ConfigProperties
	Set-Add
	Set-Logging
	Set-Display

###Examples

####Build Configuration XML:

1. Open a PowerShell console.

            From the Run dialog type PowerShell
            
2. Change the directory to the location where the PowerShell Script is saved.

            Example: cd C:\PowerShellScripts
            
3. Dot-Source the script to gain access to the functions inside.

           Type: . .\OfficeConfiguration.ps1

           By including the additional period before the relative script path you are 'Dot-Sourcing' 
           the PowerShell function in the script into your PowerShell session which will allow you to 
           run the inner functions from the console.
	
4. Run the commands (All commands are able to be piped into each other except New-OfficeConfiguration needs to be first.

            Example: 
			New-OfficeConfiguration -Bitness "64" -ProductId "O365ProPlusRetail" -OutPath "$env:Public/Documents/config.xml" | 
			Remove-Product -All | 
			Add-Product -ProductId "O365ProPlusRetail" -LanguageId ("en-US", "es-es") -ExcludeApps ("Access", "InfoPath") | 
			Set-Updates -Enabled "True" -UpdatePath "\\Server\share\" -Deadline "05/16/2014 18:30" -TargetVersion "15.1.2.3" | 
			Set-ConfigProperties -ForceAppShutDown "True" -PackageGUID "12345678-ABCD-1234-ABCD-1234567890AB" | 
			Set-Add -SourcePath "C:\Preload\Office" -Version "15.1.2.3" | 
			Set-Logging -Level "Standard" -Path "%temp%" | 
			Set-Display -Level "none" -AcceptEULA "True"
            


