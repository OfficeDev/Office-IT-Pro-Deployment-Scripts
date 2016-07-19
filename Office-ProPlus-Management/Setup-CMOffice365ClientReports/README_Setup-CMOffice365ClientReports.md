#Configure and setup Office 365 client reports in System Center Configuration Manager

##Before you begin
###Copy the required files locally
*       Configuration.txt  
*	Office365ClientConfigurations.mof  
*	Office365ProPlusConfigurations2013.mof  
*	Setup-Office365ClientReports.ps1  
*	Reports  
      *	Office 365 Client computers with shared activation.rdl  
      *	Office 365 Client count per channel.rdl  
      *	Office 365 Client languages per product.rdl  
      *	Office 365 Client product versions count with graph.rdl  
      *	Office 365 Client product versions count.rdl  
      *	Office 365 Client products count.rdl  
      *	Office 365 Client products with graph.rdl  
      *	Office 365 Client update source details with graph.rdl  
      *	Office 365 Client update source details.rdl  
      *	Office 365 Client update status for each computer.rdl  
      *	Office 365 Client users and their associated computers with details.rdl  
      
##Deactivate the current Office365ProPlusConfigurations class

Update 1606 for Configuration Manager introduced an inventory class to collect some of the Office 365 client 
information. In order to collect all of the necessary information required for the Office 365 client reports 
we should disable this class and add the two required Inventory classes; **Office365ClientConfigurations** 
and **Office365ProPlusConfigurations2013**.

1.	Open the Configuration Manager console.
2.	Go to **Administration/Overview/Client Settings**.
3.	Right click on **Default Client Settings** and choose **Properties**.
4.	Click **Hardware Inventory**.
5.	Click **Set Classes …**
6.	Deselect **OFFICE365PROPLUSCONFIGURATIONS** and click **OK**.
7.	Click **OK** to close the Default Settings window.


##Add the new classes into Hardware Inventory
###Step 1: Append the configuration.mof file  

Before we can create the reports we need to modify the configuration.mof file with the new class information. There are two options available:
      
  * PowerShell 
  
        The PowerShell script will locate the installation directory of Configuration Manager and 
        append the contents of configuration.txt into the configuration.mof file. A backup copy of 
        the existing configuration.mof file will be created in the same directory before the file is appended.  
  *	Manually  
  
        The IT Pro will need to locate the configuration.mof file and manually copy the contents 
        of configuration.txt to the end of the file.  
        
You can find more information about MOF files at https://technet.microsoft.com/en-us/library/bb632896.aspx

####Use PowerShell to copy the content to the MOF file
1.	Open a PowerShell console  

        From the Run dialog type PowerShell. Open the program as an administrator.
2.	Change the directory to the location where the PowerShell script is saved  

        Example: cd c:\PowerShellScripts
3.	Dot-Source the Edit-ConfigurationMofFile function into your current session  

        Type . .\Setup-Office365ClientReports 
      
        By Including the additional period before the relative script path you are 'Dot-sourcing' 
        the PowerShell function in the script into the PowerShell session which will allow you to run 
        the function from the console.
4.	Run the function against the local Configuration Manager server  

        Edit-ConfigurationMofFile
        
####Manually copy the content into the MOF file
1.	**Copy** the contents of **configuration.txt**.  
2.	Navigate to the Configuration Installation folder. Default installations will be inside **`<installation directory>`\Program Files\Microsoft Configuration Manager**.  
3.	Create a backup copy of configuration.mof.  
4.	**Open** the configuration.mof from **inboxes\clifiles.src\hinv\configuration.mof**.  
5.	**Paste the contents** from configuration.txt to the **end** of the configuration.mof file.  
6.	**Save** and **close** configuration.mof.  

###Step 2: Enable the new class in Configuration Manager
1.	From the Configuration Manager Console go to **Administration**, expand **Site Configuration**, then click **Client Settings**.  
2.	Right click **Default Client Settings** and choose **Properties**.  
3.	Click on **Hardware Inventory**, then click **Set Classes**.  
4.	Click on **Import**, navigate to the **Office365ClientConfigurations.mof** file and click **Open**.  
5.	On the Import Summary window click **Import**.  
6.	Repeat steps 4 - 5 and choose **Office365ProPlusConfigurations2013.mof**.  
7.	In the Hardware Inventory Classes window the new Office365ClientConfigurations and Office365ProPlusConfigurations2013 classes will be at the top of the list. You can leave the check mark to apply the classes to all devices, or uncheck the classes to enable on a custom Client Device Setting.   
8.	Click **OK**.  

###Step 3: Enable the classes in a custom Client Device Setting

If you enabled the new classes in Default Client Settings you can move on to Step 5. If you want to enable the new class on a custom Client Device Setting proceed with the following steps.  

1.	From **Administration/Site Configuration/Client Settings** right click on the **Device Setting** you want to enable the new hardware inventory classes on and choose **Properties**.  
2.	Click **Hardware Inventory** and select **Set Classes**.  
3.	Scroll down until you see **Office365ClientConfigurations** and **Office365ProPlusConfigurations2013**, **check the boxes**, and click **OK**.  
4.	Click **OK** on the Device Settings window.  

###Step 4: Deploy the Device Setting to a collection
1.	**Right click** on the **Device Setting** that has the Office365ClientConfigurations classes enabled and choose **Deploy**.  
2.	Select the **collection** and click **OK**.  

**Note** - When the next hardware inventory runs on the client the information from the new class will be collected. This may take some time.
          
##Import the custom reports

The Office 365 Client reports can be imported using PowerShell or by manually uploading the files to the reporting server.

After the Office 365 Client reports are imported they will be available in the Configuration Manager console. The reports will 
start to show data once the hardware inventory has run on the clients. Completion time will depend on the size of the environment and the frequency of hardware inventory scans.

###Use PowerShell to import the reports
1.	Open a PowerShell console.  

          From the Run dialog type PowerShell. Open the program as an administrator.
2.	Change the directory to the location where the PowerShell script is saved.  

          Example: cd c:\PowerShellScripts
3.	Dot-Source the Import-CMReports function into your current session.  

          Type . .\Setup-Office365ClientReports

          By including the additional period before the relative script path you are 'Dot-Sourcing' 
          the PowerShell function in the script into your PowerShell session. This allows you to 
          run the function from the console.

4.	Run the function against the local Report Server.  

          Example: Import-CMReports  

          This will run the script and use the local computer name as the Report Server.  
          A folder called Custom Reports will be created and used to host the report files on the report server.  
          
          Example: Import-CMReports –FolderName “Software – Office 365 Clients”  
	  
          This will run the script and use the local computer name as the Report Server.
          A folder called Software – Office 365 Clients will be created and used to host the report files on the report server.

          
###Manually upload the reports
1.	From **Report Manager**, navigate to the Contents page.  
2.	Click **Upload File**.  
3.	Click **Browse**.  
4.	Select one of the client rdl files.  
5.	If you want to replace an existing item with the new item, select Overwrite item if it exists.  
6.	Click **OK**.








