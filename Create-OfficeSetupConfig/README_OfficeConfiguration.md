#Office Configuration XML Builder

Functions for building the configuration xml file for click to run office products
Functions included:
	New-ODTConfiguration
	Add-ODTProduct
	Remove-ODTProduct
	Set-ODTUpdates
	Set-ODTConfigProperties
	Set-ODTAdd
	Set-ODTLogging
	Set-ODTDisplay

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
			New-ODTConfiguration -Bitness "64" -ProductId "O365ProPlusRetail" -OutPath "$env:Public/Documents/config.xml" | 
			Remove-ODTProduct -All | 
			Add-ODTProduct -ProductId "O365ProPlusRetail" -LanguageId ("en-US", "es-es") -ExcludeApps ("Access", "InfoPath") | 
			Set-ODTUpdates -Enabled "True" -UpdatePath "\\Server\share\" -Deadline "05/16/2014 18:30" -TargetVersion "15.1.2.3" | 
			Set-ODTConfigProperties -ForceAppShutDown "True" -PackageGUID "12345678-ABCD-1234-ABCD-1234567890AB" | 
			Set-ODTAdd -SourcePath "C:\Preload\Office" -Version "15.1.2.3" | 
			Set-ODTLogging -Level "Standard" -Path "%temp%" | 
			Set-ODTDisplay -Level "none" -AcceptEULA "True"
            


