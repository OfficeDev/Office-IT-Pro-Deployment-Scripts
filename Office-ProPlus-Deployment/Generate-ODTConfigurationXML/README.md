# **Generate Office Deployment Tool (ODT) Configuration XML**

This PowerShell function queries the existing configuration of the target computer and generates the Configuration XML for Click-to-Run for Office 365 products.  This XML is used with the [Office Deployment Tool (ODT)](http://www.microsoft.com/en-us/download/details.aspx?id=36778) to deploy Office Click-To-Run products.  The purpose of this script is to dynamically generate a configuration.xml file to be used to either install new or modify existing Office Click-To-Run deployments.  

Deploying Office can be challenging in Organizations that have to support many different languages.  This script provides a way to automate the deployment of Language packs.  The script will query for the languages currently in use by the local computer.  It will then add those languages into the configuration xml.

You can control which languages the script will add to the configuration xml by using the **Languages** parameter. The parameter has four options.  The options and explanations for this parameter are listed below.

 - **CurrentOfficeLanguages** - This option will use the languages that the current installation of Office is using.
 - **OSLanguage** - This option will use only the primary language of the Operating System.
 - **OSandUserLanguages** - This option will use use the primary language of the Operating System and the languages that the local users have added to their profiles.
 - **AllInUseLanguages** - This option will use all of the currently in use languages including, Office, Operating System and user lanaguages.

For more information on the specifics of the Click-to-Run for Office 365 Configuration XML go to the following link.
[Click-to-Run for Office 365 Configuration XML Reference](https://technet.microsoft.com/en-us/library/JJ219426.aspx)

If the parameter **IncludeUpdatePathAsSourcePath** is set to $true then it will use the UpdatePath as the SourcePath for the generated configuration xml. This option would be useful for distributed environments where clients are pointed to a local update source for updates.  In order to use this option the Update source must have the version and language packs that are required by the generated configuration xml.  The UpdatePath is stored in the UpdateUrl value in the registy path **HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration** for Office 2013 ProPlus or **HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration** for Office 2016 ProPlus.

### **Examples**

1. Open a PowerShell console.

		From the Run dialog type PowerShell 

2. Change directory to the location where the PowerShell Script is saved.

		Example: cd C:\PowerShellScripts

3. Dot-Source the Generate-ODTConfigurationXML function into your current session.

		Type . .\Generate-ODTConfigurationXML.ps1
		By including the additional period before the relative script path you are 'Dot-Sourcing' 
		the PowerShell function in the script into your PowerShell session which will allow you to 
		run the function from the console.

4. Run the function against the local computer

		Generate-ODTConfigurationXml -Languages AllInUseLanguages -TargetFilePath configuration.xml 

5. An example output is below.  The first language in the list is the Shell UI culture.  

          <Configuration>
             <Add Version="16.0.4745.1002" OfficeClientEdition="32" Channel="Current">
                 <Product ID="O365ProPlusRetail">
                   <Language ID="en-us" />
                   <Language ID="de-de" />
                   <Language ID="fr-fr" />
                 </Product>
                 <Product ID="ProjectProRetail">
                   <Language ID="en-us" />
                   <Language ID="de-de" />
                   <Language ID="fr-fr" />
                 </Product>
             </Add>
             <Updates Enabled="False" />
          </Configuration>

6. In order to create a complete solution there are two other scripts in this GitHub Repository.  

	[Edit-OfficeConfigurationFile](../Edit-OfficeConfigurationFile) - This script provides functions to edit the Office Click-To-Run configuration file.
	
	[Install-OfficeClickToRun](../Install-OfficeClickToRun) - This script will install Office Click-To-Run.  The script requires that a configuration Xml file is provided.
	
[![Analytics](https://ga-beacon.appspot.com/UA-70271323-4/README_Generate-ODTConfigurationXML?pixel)](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts)
