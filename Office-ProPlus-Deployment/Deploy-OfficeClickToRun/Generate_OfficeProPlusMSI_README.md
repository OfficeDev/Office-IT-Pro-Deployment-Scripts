#Package an Office 365 ProPlus installation

The Install Toolkit is an application that will package an Office 365 ProPlus installation into a single Executable or Windows Installer Package (MSI) file. 
The XML configuration file is embedded in the file which allows you to easily distribute Office 365 ProPlus with a custom configuration. 
The Install Tookkit is a desktop application that is installed via Microsoft ClickOnce. Please note that the Install Toolkit requires .Net 4.5 in order to run properly. 

1. Go to the Configuration XML Editor website located [here](http://officedev.github.io/Office-IT-Pro-Deployment-Scripts/XmlEditor.html).  
2. In the left panel under Tools click on **Install Toolkit**.  
3. Click **Launch Installation**.  
4. Run the **OfficeProPlusInstallGenerator.application** file.  
5. Click **Install** on the Application Install – Security Warning prompt. Accept the security warnings.  
6. The Install Toolkit will open automatically. To begin the process of packaging a deployment file make sure Create new Office 365 ProPlus installer is selected and click Start.  
7. Select the main Office product, Office 365 ProPlus or Office 365 for Business.  
8. Select the edition of Office to install, 32-Bit or 64-Bit.  
9. Click the dropdown under Channel and choose the channel you would like to deploy.  
10. Check the box next to any of the additional Office products you need.  
11. Click **Next**.  
12. Click Add Language and choose any additional languages you may need to install and click OK. If a language other than English needs to be the primary highlight the necessary language and click Set Primary.  
13. Click **Next**.  
14. Use the default Version to install the latest version.  
15. Add a Remote Logging Path, Source Path, or Download Path and click Next.  
16. Deselect any applications that need to be excluded from the deployment and click Next.  
17. Click **Next** in the Optional window.  
18. Updates are enabled by default. Deselect Enable to turn Updates off and continue to step 21.  
19. Verify the Channel is the same as the channel select in step 9.  
20. You may add the Update Path, Target Version and Deadline or leave them as default.  
21. Click **Next**.  
22. Select or deselect the list of additional available options and click **Next**.  
23. Choose MSI or Executable.  
24. If you need to sign the installer using a certificate check the box next to Sign installer and click **Select Certificate** or **Generate Certificate**.  
25. Add a version or leave as default.  
26. Choose Silent install to run the file silently.  
27. Enter the file path to save the generated file or leave as default.  
28. Check Embed Office installation file to include the Office files with the MSI or Exe. Leave this check box blank if you have chosen to use a local Source Path or to install from the Microsoft content delivery network (CDN).  
29. Click **Generate** and click **OK** when the process has finished.  