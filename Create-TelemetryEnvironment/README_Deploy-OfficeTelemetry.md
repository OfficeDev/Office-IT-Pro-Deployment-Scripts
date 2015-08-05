#Deploy Office Telemetry

Configure Office Telemetry Dashboard. If SQL Server is not installed SQL Server 2014 Express 
will be installed. A database will be set up using the built-in settings for Office Telemetry.
A shared folder will be created and permissions will be set up to allow telemetry agents to 
upload data.

###Pre-requisites

The Deploy-TelemetryDashboard script must be ran from a machine with Office 2013 or 2016 already installed.

Copy the OfficeTelemetryDatabase.sql file to C:\Users\username\Appdata\Local\Temp where username
is the name of the user logged in.

The user logged in must have administrative privelages and PowerShell needs to be opened as an administrator.

The 2013 or 2016 administrative templates need to be installed on the Domain Controller.

.NET Framework 3.5 must be installed. If it is not enabled the script will enable it.

Links:

2013 Administrative Templates: https://www.microsoft.com/en-us/download/details.aspx?id=35554

Overview of Office Telemetry: https://technet.microsoft.com/en-us/library/JJ863580.aspx

SQL Server 2014 Express download: https://www.microsoft.com/en-us/download/details.aspx?id=42299

###Example

####Install SQL, configure a database, install the telemetry processor, and enable the agent to upload data

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

1. From the Domain Controller open a PowerShell console.

          From the Run dialog type PowerShell
          
2. Change the directory to the location where the PowerShell Script is saved.

          Example: cd C:\PowerShellScripts
          
3. Run the script.

          Type . .\TelemetryGpo.ps1
          
          By including the additional period before the relative script path you are 'Dot-Sourcing' 
          the PowerShell function in the script into your PowerShell session which will allow you to 
          run the function 'Get-ModernOfficeApps' from the console.
          
4. Type the SQL server and press Enter.

5. Type the version of Microsoft Office and press Enter.
