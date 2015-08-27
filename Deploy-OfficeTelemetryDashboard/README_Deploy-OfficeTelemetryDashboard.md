#Deploy Office Telemetry

This script will configure the Office Telemetry Dashboard in your environment. If SQL Server is not installed SQL Server 2014 Express will be installed. A database will be set up using the standard settings for Office Telemetry found in the dpconfig.exe file.
A shared folder will be created and configured to allow telemetry agents to upload data. A Group Policy can be created to enable telemetry agents on computers in a domain. Computers with versions of Office older than 2013 will need to have the telemetry agent installed. Follow the instructions to create the GPO that will install and enable the telemetry agent on computers with versions of Office older than 2013.

###Pre-requisites

1. The Deploy-TelemetryDashboard.ps1 script must be ran from a machine with **Office 2013** or **Office 2016**.  The server that you run this script from will be the Telemetry Processing server where clients will submit Telemetry.

2. The user logged in must have administrative privelages and PowerShell needs to be opened as an administrator.

3. .NET Framework 3.5 must be installed. If it is not enabled the script will enable it.

###Links:

Overview of Office Telemetry: https://technet.microsoft.com/en-us/library/JJ863580.aspx

2013 Administrative Templates: https://www.microsoft.com/en-us/download/details.aspx?id=35554

SQL Server 2014 Express download: https://www.microsoft.com/en-us/download/details.aspx?id=42299

.NET Framework 3.5 download: https://www.microsoft.com/en-us/download/details.aspx?id=21

###Setup

1. All of the files below must be in the directory from where you run the scripts.

          Deploy-TelemetryDashboard.ps1
          Configure-TelemetryGpo.ps1
          Deploy-TelemetryAgent.ps1
          OfficeTelemetryDatabase.sql
          osmia32.msi
          osmia64.msi

###Examples

#####If SQL server is not found on the local server SQL Server 2014 Express will be downloaded and installed, a database will be configured using the OfficeTelemetryDatabase.sql file, two shared folders will be created and configured, the Telemetry Processor will be installed, and the Telemetry Agent will be enabled to collect and upload data.

1. Open a PowerShell console.

          From the Run dialog type PowerShell, right click, and choose Run as Administrator
            
2. Change the directory to the location where the PowerShell Script is saved.

          Example: cd C:\PowerShellScripts
            
3. Run the Script.

          .\Deploy-TelemetryDashboard.ps1

4. Wait for the script to finish. When the script is completed restart the computer to allow the 
telemetry agent scheduled task to run and collect data.

####Create a GPO on the Domain Controller.

A Group Policy can be set to enable Telemetry Agent uploading and logging on computers in the domain.

1. From the Domain Controller open a PowerShell console as an administrator.

          From the Run dialog type PowerShell, right click, and choose Run as Administrator.
          
2. Change the directory to the location where the PowerShell Script is saved.

          Example: cd C:\PowerShellScripts
   
3. To Create or Configure a Group Policy to configure Office clients to submit Telemetry to the Telemetry proccessing server run the following Powershell Script.

          .\Configure-TelemetryGpo.ps1 -GpoName "Office Telemetry" -TelemetryServer "TelemetryServerName"

6. From a computer in the OU open a command prompt and type gpupdate /force and press Enter.

7. Restart the computer, log in, and wait for the script to run in the background.
