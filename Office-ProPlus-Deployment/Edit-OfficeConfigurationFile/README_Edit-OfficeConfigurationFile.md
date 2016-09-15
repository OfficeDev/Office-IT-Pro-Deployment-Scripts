### Edit Office Configuration File
Script to modify the Configuration.xml file to configure installation options for Click-to-Run for Office 365 products.

**IT Pro Scenario:** For organizations that are upgrading Office and want to control the customizations of the deployment. This is best used for organizations that want to ensure certain functions are included within the xml file, but don't want to create it from scratch. Example, if everyone in the organization needs visio, you can use this script to update all xml files to include visio in the installation. 

[README](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/wiki/README_Office-ConfigurationXMLBuilder)

1. Open a PowerShell console.

		From the Run dialog type PowerShell 

2. Change directory to the location where the PowerShell Script is saved.

		Example: cd C:\PowerShellScripts

3. Dot-Source the Edit-OfficeConfigurationFile function into your current session.

		Type . .\Edit-OfficeConfigurationFile.ps1

		By including the additional period before the relative script path you are 'Dot-Sourcing' 
		the PowerShell function in the script into your PowerShell session which will allow you to 
		run the function from the console.

4 Run the Edit-OfficeConfigurationFile functions to edit the configuration xml

Function        Add-ODTProductToAdd                                                                                                                                                                                                                              
Function        Add-ODTProductToRemove
Function        Get-LanguagesFromXML                                                                                                                                                                                                                             
Function        Get-ODTAdd                                                                                                                                                                                                                                       
Function        Get-ODTConfigProperties                                                                                                                                                                                                                          
Function        Get-ODTDisplay                                                                                                                                                                                                                                   
Function        Get-ODTLogging                                                                                                                                                                                                                                   
Function        Get-ODTProductToAdd                                                                                                                                                                                                                              
Function        Get-ODTProductToRemove                                                                                                                                                                                                                           
Function        Get-ODTUpdates  
Function        New-ODTConfiguration    
Function        Remove-ODTAdd                                                                                                                                                                                                                                    
Function        Remove-ODTConfigProperties                                                                                                                                                                                                                       
Function        Remove-ODTDisplay                                                                                                                                                                                                                                
Function        Remove-ODTExcludeApp                                                                                                                                                                                                                             
Function        Remove-ODTLogging                                                                                                                                                                                                                                
Function        Remove-ODTProductToAdd                                                                                                                                                                                                                           
Function        Remove-ODTProductToRemove                                                                                                                                                                                                                        
Function        Remove-ODTUpdates   
Function        Set-ODTAdd                                                                                                                                                                                                                                       
Function        Set-ODTConfigProperties                                                                                                                                                                                                                          
Function        Set-ODTDisplay                                                                                                                                                                                                                                   
Function        Set-ODTLogging                                                                                                                                                                                                                                   
Function        Set-ODTProductToAdd                                                                                                                                                                                                                              
Function        Set-ODTUpdates                                                                                                                                                                                                                                   
Function        Show-ODTConfiguration   
Function        Test-UpdateSource  
Function        Undo-ODTLastChange  
Function        Validate-UpdateSource                                                                                                                                                                                                                        
