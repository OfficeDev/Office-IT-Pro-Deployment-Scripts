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

1. Copy the below files in to the folder to where the script will be ran.

          Create-TelemetryGpo.ps1
          Deploy-TelemetryDashboard.ps1
          
2. Copy the OfficeTelemetryDatabase.sql file to C:\Users\username\Appdata\Local\Temp (%temp%) where username
is the name of the user logged in.

          This file contains the predefined database settings found in the dpconfig.exe file.
          
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

A Group Policy can be set to enable Telemetry Agent uploading and logging on computers in the domain.

1. From the Domain Controller open a PowerShell console as an administrator.

          From the Run dialog type PowerShell, right click, and choose Run as Administrator.
          
2. Change the directory to the location where the PowerShell Script is saved.

          Example: cd C:\PowerShellScripts
          
3. Run the script and specify the name of the GPO.

          Type . .\Create-TelemetryGpo -GpoName "Office Telemetry"
          
          By including the additional period before the relative script path you are 'Dot-Sourcing' 
          the PowerShell function in the script into your PowerShell session which will allow you to 
          run the function 'Get-ModernOfficeApps' from the console.

4. Run run the script to create the GPO and set the registry values to enable Telemetry Agent logging and uploading.

          Type . .\Create-TelemetryGpo -GpoName "Office Telemetry" -CommonFileShare "Server1" -officeVersion 2013
          
          By including the additional period before the relative script path you are 'Dot-Sourcing' 
          the PowerShell function in the script into your PowerShell session which will allow you to 
          run the function 'Get-ModernOfficeApps' from the console.

####Configure the GPO to run on startup

1. From the Domain Controller open a PowerShell console as an administrator.

          From the Run dialog type PowerShell, right click, and choose Run as Administrator.
          
2. Change the directory to the location where the PowerShell Script is saved.

          Example: cd C:\PowerShellScripts
          
3. Run the script, specify the GPO name and the shared drive the telemetry agent will upload data to.

          Type . .\Set-TelemetryStartup -GpoName "Office Telemetry" -CommonFileShare "\\Server1\TDShared"
          
          By including the additional period before the relative script path you are 'Dot-Sourcing' 
          the PowerShell function in the script into your PowerShell session which will allow you to 
          run the function 'Get-ModernOfficeApps' from the console.

4. Link the GPO to the correct OU in Group Policy Management.

	  1. Right click on the correct OU and choose Link an existing GPO...

	  2. Highlight the GPO and click OK.

5. From a computer in the OU open a command prompt and type gpupdate /force and press Enter.

6. Restart the computer, log in, and wait for the script to run in the background.
