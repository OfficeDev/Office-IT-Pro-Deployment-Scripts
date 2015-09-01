#Office Configuration XML Builder

The Click-to-Run for Office 365 Configuration.xml file is used to specify Click-to-Run installation and update options. The Office Deployment Tool is a downloadable tool that includes a sample Configuration.xml file. Administrators can modify the Configuration.xml file to configure installation options for Click-to-Run for Office 365 products.

*Reference: https://technet.microsoft.com/en-us/library/JJ219426.aspx*

This script includes functions to create and edit the configuration xml file. It is designed to reduce the effort required to create or makes changes to this configuration file while also reducing the chance of formatting errors in xml.

Functions included:

	Add-ODTProductToAdd
	Add-ODTProductToRemove
	Get-ODTAdd
	Get-ODTConfigProperties
	Get-ODTDisplay
	Get-ODTLogging
	Get-ODTProductToAdd
	Get-ODTProductToRemove
	Get-ODTUpdates
	New-ODTConfiguration
	Remove-ODTAdd
	Remove-ODTConfigProperties
	Remove-ODTDisplay
	Remove-ODTLogging
	Remove-ODTProductToAdd
	Remove-ODTProductToRemove
	Remove-ODTUpdates
	Set-ODTAdd
	Set-ODTConfigProperties
	Set-ODTDisplay
	Set-ODTLogging
	Set-ODTUpdates
	Show-ODTConfiguration
	Undo-ODTLastChange

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
		New-ODTConfiguration -Bitness "64" -ProductId "O365ProPlusRetail" -LanguageId "en-us" -TargetFilePath "$env:UserProfile\Desktop\configuration.xml" |
		Remove-ODTProductToAdd -All | 
		Add-ODTProductToAdd -ProductId "O365ProPlusRetail" -LanguageIds ("en-us") -ExcludeApps ("Access", "InfoPath") | 
		Set-ODTUpdates -Enabled "True" -UpdatePath "\\Server\share\" -Deadline "05/16/2014 18:30" -TargetVersion "15.1.2.3" | 
		Set-ODTConfigProperties -ForceAppShutDown "True" -PackageGUID "12345678-ABCD-1234-ABCD-1234567890AB" |
		Set-ODTAdd -SourcePath "C:\Preload\Office" -Version "15.1.2.3" | 
		Set-ODTLogging -Level "Standard" -Path "%temp%" | 
		Set-ODTDisplay -Level "none" -AcceptEULA "True"


