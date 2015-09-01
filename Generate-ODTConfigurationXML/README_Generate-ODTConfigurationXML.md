#**Generate Office Deployment Tool Configuration XML**

This PowerShell function queries the existing configuration of the target computer and generates the Configuration XML for Click-to-Run for Office 365 products.  This XML is used with the [Office Deployment Tool (ODT)](http://www.microsoft.com/en-us/download/details.aspx?id=36778) to deploy Office Click-To-Run products.  The purpose of this script is to dynamically generate a configuration.xml file to be used to either install new or modify existing Office Click-To-Run deployments.  

Deploying Office can be challenging in Organizations that have to support many different languages.  This script provides a way to automate the deployment of Language packs.  The script will query for the languages currently in use by the local computer.  It will then add those languages into the configuration xml.

You can control which languages the script will add to the configuration xml by using the **Languages** parameter. The parameter has four options.  The options and explanations for this parameter are listed below.

 - **CurrentOfficeLanguages** - This option will use the languages that the current installation of Office is using.
 - **OSLanguage** - This option will use only the primary language of the Operating System.
 - **OSandUserLanguages** - This option will use use the primary language of the Operating System and the languages that the local users have added to their profiles.
 - **AllInUseLanguages** - This option will use all of the currently in use languages including, Office, Operating System and user lanaguages.

For more information on the specifics of the Click-to-Run for Office 365 Configuration XML go to the following link.
[Click-to-Run for Office 365 Configuration XML Reference](https://technet.microsoft.com/en-us/library/JJ219426.aspx)

###**Examples**

1. Open a PowerShell console.

		From the Run dialog type PowerShell 
		
2. Change directory to the location where the PowerShell Script is saved.

		Example: cd C:\PowerShellScripts
		
2. Run the Script. With no parameters specified the script will return the locally installed Office Version.

		Type . .\Generate-ODTConfigurationXML.ps1
		Press Enter and then if Microsoft Office is installed locally it should display. 
		By including the additional period before the relative script path you are 'Dot-Sourcing' 
		the PowerShell function in the script into your PowerShell session which will allow you to 
		run the function from the console.
	
	

	

