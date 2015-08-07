#Deploy Office Telemetry

Configure the Office Telemetry Dashboard. If SQL Server is not installed SQL Server 2014 Express 
will be installed. A database will be set up using the standard settings for Office Telemetry found in the dpconfig.exe file.
A shared folder will be created and configured to allow telemetry agents to upload data. A Group Policy can be created to enable telemetry agents on computers in a domain. Computers with versions of Office older than 2013 will need to have the telemetry agent installed. Follow the instructions to create the GPO that will install and enable the telemetry agent on computers with versions of Office older than 2013.

###Pre-requisites

1. The Deploy-TelemetryDashboard.ps1 script must be ran from a machine with Office 2013.

2. The user logged in must have administrative privelages and PowerShell needs to be opened as an administrator.

3. .NET Framework 3.5 must be installed. If it is not enabled the script will enable it.

###Links:

Overview of Office Telemetry: https://technet.microsoft.com/en-us/library/JJ863580.aspx

2013 Administrative Templates: https://www.microsoft.com/en-us/download/details.aspx?id=35554

SQL Server 2014 Express download: https://www.microsoft.com/en-us/download/details.aspx?id=42299

.NET Framework 3.5 download: https://www.microsoft.com/en-us/download/details.aspx?id=21

###Setup

1. Copy the below files in to the folder from where the script will be ran.

          Create-TelemetryGpo.ps1
          Deploy-TelemetryAgent.ps1
          Deploy-TelemetryDashboard.ps1
          Set-TelemetryStartup.ps1
          
2. Copy the OfficeTelemetryDatabase.sql file to C:\Users\username\Appdata\Local\Temp (%temp%) where username
is the name of the user logged in.
          
3. Before creating the GPO, if you are testing on computers in the domain with Office versions older than 2013 copy the osmia32 and osmia64 msi files in to a shared folder. 

###Examples

####Check for SQL installations, if a SQL server is not found SQL Server 2014 Express will be downloaded and installed, a database will be configured using the OfficeTelemetryDatabase.sql file, a shared folder will be created and configured, the Telemetry Processor will be installed, and the Telemetry Agent will be enabled to collect and upload data.

1. Open a PowerShell console.

          From the Run dialog type PowerShell, right click, and choose Run as Administrator
            
2. Change the directory to the location where the PowerShell Script is saved.

          Example: cd C:\PowerShellScripts
            
3. Run the Script.

          Type . .\Deploy-OfficeTelemetry.ps1
          
          By including the additional period before the relative script path you are 'Dot-Sourcing' 
          the PowerShell function in the script into your PowerShell session which will allow you to 
          run the function 'Get-ModernOfficeApps' from the console.
          
4. Wait for the script to finish. When the script is completed restart the computer to allow the 
telemetry agent scheduled task to run and collect data.

####Create a GPO on the Domain Controller

A Group Policy can be set to enable Telemetry Agent uploading and logging on computers in the domain. If computers in
the domain have versions of Office older than 2013 only the GPO will be created.

1. From the Domain Controller open a PowerShell console as an administrator.

          From the Run dialog type PowerShell, right click, and choose Run as Administrator.
          
2. Change the directory to the location where the PowerShell Script is saved.

          Example: cd C:\PowerShellScripts
          
3. For computers with versions of Office older than 2010 run the script to create the GPO. Specify the name of the GPO and the name of the SQL Server.

          Type . .\Create-TelemetryGpo -GpoName "Office Telemetry" -SqlServerName SQLExpress
          
          By including the additional period before the relative script path you are 'Dot-Sourcing' 
          the PowerShell function in the script into your PowerShell session which will allow you to 
          run the function 'Get-ModernOfficeApps' from the console.

4. For computers with versions of Office newer than 2010 run the script to create the GPO and set the registry values to enable Telemetry Agent logging and uploading.

          Type . .\Create-TelemetryGpo -GpoName "Office Telemetry" -SqlServerName SQLExpress -officeVersion 2013

####Modify the GPO to copy the Deploy-TelemetryAgent.ps1 script to the Startup folder. 

Computers in the domain with versions of Office older than 2013 will copy the osmia32.msi or osmia64.msi file, depending on the computer's bitness, to the temp folder (%temp%) and will install. The script will also create the registry keys and values needed for the telemetry agent to collect and upload data to the telemetry shared folder.

1. From the Domain Controller open a PowerShell console as an administrator.

          From the Run dialog type PowerShell, right click, and choose Run as Administrator.
          
2. Change the directory to the location where the PowerShell Script is saved.

          Example: cd C:\PowerShellScripts
          
3. Run the script. Specify the GPO name, UNC path of the shared folder hosting the osmia32 and osmia64 msi files, and specify the shared folder created in the Deploy-TelemetryDashboard.ps1 script where the telemetry data is uploaded.

          Type . .\Set-TelemetryStartup -GpoName "Office Telemetry" -UncPath "\\Server1\Sharedfolder" -CommonFileShare "\\Server1\TDShared"
          
          By including the additional period before the relative script path you are 'Dot-Sourcing' 
          the PowerShell function in the script into your PowerShell session which will allow you to 
          run the function 'Get-ModernOfficeApps' from the console.
          
4. To verify the Deploy-TelemetryAgent.ps1 script is in the GPO startup folder:

          1. Open Group Policy Management
          2. Right click the newly created GPO and choose Edit.
          3. Under Computer Configuration click the drop down next to Policies.
          4. Click the drop down next to Windows Setting and click on Scripts.
          5. Double click Startup and click the PowerShell Scripts tab.
          6. Verify Deploy-TelemetryAgent.ps1 is under Name and -UncPath \\sharedfolder\path is under Parameters where \\sharedfolder\path is the path you entered.
          7. Click OK to close the Startup Properties window.
          
5. Link the new GPO to the correct OU in the domain.

          1. Right click on the apporpriate OU and choose Link an Existing GPO...
          2. Select the GPO click OK.

6. Refresh the group policy from a computer in the domain and restart.

          1. From a run dialogue box type cmd and press Enter.
          2. Type gpupdate -force and press Enter.
          3. Restart the computer.
          
