# **Generate Office Deployment Tool (ODT) LanguagePack XML**

This PowerShell function will create a new XML file that can be used to deploy additional languages packs to a computer with Office 365 ProPlus already installed. This XML is used with the [Office Deployment Tool (ODT)](http://www.microsoft.com/en-us/download/details.aspx?id=36778) to deploy Office Click-To-Run products.    

Deploying Office language packs can be challenging in organizations that have to support many different languages.  This script provides a way to automate the deployment of Language Packs.  

You can control the bit and language IDs that are needed for the deployment.

For more information on the specifics of deploying Office 365 ProPlus language packs go to the following link.
[Add languages to existing installations of Office 365 ProPlus](https://technet.microsoft.com/en-us/library/jj219422.aspx#Anchor_7)

### **Examples**

1. Open a PowerShell console.

		From the Run dialog type PowerShell 

2. Change directory to the location where the PowerShell Script is saved.

		Example: cd C:\PowerShellScripts

3. Dot-Source the Generate-ODTLanguagePackXML function into your current session.

		Type . .\Generate-ODTLanguagePackXML.ps1
		By including the additional period before the relative script path you are 'Dot-Sourcing' 
		the PowerShell function in the script into your PowerShell session which will allow you to 
		run the function from the console.

4. Run the function

		Generate-ODTLanguagePackXml -Languages es-es,de-de,fr-fr -TargetFilePath LanguagePacks.xml 

5. An example output is below.  

          <Configuration>
             <Add OfficeClientEdition="32">
                 <Product ID="LanguagePack">
                   <Language ID="es-es" />
                   <Language ID="de-de" />
                   <Language ID="fr-fr" />
                 </Product>
             </Add>
          </Configuration>
