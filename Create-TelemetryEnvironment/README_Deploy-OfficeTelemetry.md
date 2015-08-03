#Deploy Office Telemetry

Configure Office Telemetry Dashboard. If SQL Server is not installed SQL Server 2014 Express 
will be installed. A database will be set up using the built-in settings for Office Telemetry.
A shared folder will be created and permissions will be set up to allow telemetry agents to 
upload data.

###Pre-requisites

The script must be ran from a machine with Excel 2013 or 2016 already installed.

Copy the OfficeTelemetryDatabase.sql file to C:\Users\username\Appdata\Local\Temp where username
is the name of the user logged in.

The user logged in must have administrative privelages and PowerShell needs to be opened as an administrator.

The 2013 or 2016 administrative templates need to be installed on the Domain Controller.

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

          Type . .\Deploy-OfficeTelemetry.ps1
          
          By including the additional period before the relative script path you are 'Dot-Sourcing' 
          the PowerShell function in the script into your PowerShell session which will allow you to 
          run the function 'Get-ModernOfficeApps' from the console.
          
4. Type the SQL server and press Enter.

5. Type the version of Microsoft Office and press Enter.
