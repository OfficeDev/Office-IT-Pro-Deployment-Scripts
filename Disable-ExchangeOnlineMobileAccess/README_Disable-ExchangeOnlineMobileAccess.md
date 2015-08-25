#**Disable Exchange Online Mobile Access**

This PowerShell Function connect to Exchange Online and disable Mobile access for all of the Mailboxes in the Organization.   

###**Pre-requisites**

1. An Office 365 account that in the Exchange Admin role 'Recipient Management' or 'Organization Managment'

###**Examples**

1. Open a PowerShell console.

		From the Run dialog type PowerShell 
		
2. Change directory to the location where the PowerShell Script is saved.

		Example: cd C:\PowerShellScripts
		
2. Run the Script. With no parameters specified the script will run and prompt you for your Office 365 Admin Credentials

		Type . .\Disable-ExchangeOnlineMobileAccess
		Press Enter and then if Microsoft Office is installed locally it should display. 
		By including the additional period before the relative script path you are 'Dot-Sourcing' 
		the PowerShell function in the script into your PowerShell session which will allow you to 
		run the function from the console.
	
3. Run the Script 

		Type $credentials = Get-Credential
		Type Disable-ExchangeOnlineMobileAccess -Credentials $credentials

4. Run the Script without prompting

		$userName = "admin@tenant.onmicrosoft.com"
		$securedPassword = ConvertTo-SecureString "PlainTextPassword" -AsPlainText -Force
		$credentials = New-Object System.Management.Automation.PSCredential ($userName, $securedPassword)

		Disable-ExchangeOnlineMobileAccess -Credentials $credentials
	

	

