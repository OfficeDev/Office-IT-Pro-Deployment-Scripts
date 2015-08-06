#Deploy Office Telemetry

Configure Office Telemetry Dashboard. If SQL Server is not installed SQL Server 2014 Express 
will be installed. A database will be set up using the built-in settings for Office Telemetry.
A shared folder will be created and permissions will be set up to allow telemetry agents to 
upload data.

###Pre-requisites

1. The Deploy-TelemetryDashboard.ps1 script must be ran from a machine with Office 2013.

2. The user logged in must have administrative privelages and PowerShell needs to be opened as an administrator.

3. .NET Framework 3.5 must be installed. If it is not enabled the script will enable it.

###Links:

2013 Administrative Templates: https://www.microsoft.com/en-us/download/details.aspx?id=35554

Overview of Office Telemetry: https://technet.microsoft.com/en-us/library/JJ863580.aspx

SQL Server 2014 Express download: https://www.microsoft.com/en-us/download/details.aspx?id=42299

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

A Group Policy can be set to enable Agent uploading and logging on computers in the domain. If computers in
the domain have Office versions older than 2013 only the GPO will be created.

1. From the Domain Controller open a PowerShell console.

          From the Run dialog type PowerShell
          
2. Change the directory to the location where the PowerShell Script is saved.

          Example: cd C:\PowerShellScripts
          
3. Run the script to create the GPO.

          Type . .\Create-TelemetryGpo -GpoName "Office Telemetry" -SqlServerName SQLExpress
          
          By including the additional period before the relative script path you are 'Dot-Sourcing' 
          the PowerShell function in the script into your PowerShell session which will allow you to 
          run the function 'Get-ModernOfficeApps' from the console.

4. Run the script to create the GPO and set the registry values for Office 2013.

          Type . .\Create-TelemetryGpo -GpoName "Office Telemetry" -SqlServerName SQLExpress -officeVersion 2013

####Modify the GPO to copy the Deploy-TelemetryAgent.ps1 script to the Startup folder. 

Computers on the domain with Office versions older than 2013 will copy the osmia32.msi or osmia64.msi file, depending on the computer's bitness, to the temp folder (%temp%) and will install. The script will also create the registry keys and values needed for the telemetry agent to collect and upload data to the telemetry shared folder.

1. From the Domain Controller open a PowerShell console as an administrator.

          From the Run dialog type PowerShell, right click, and choose Run as Administrator.
          
2. Change the directory to the location where the PowerShell Script is saved.

          Example: cd C:\PowerShellScripts
          
3. Run the script. Specify the GPO name and UNC path of the shared folder.

          Type . .\Set-TelemetryStartup -GpoName "Office Telemetry" -UncPath "\\Server1\Sharedfolder"
          
          By including the additional period before the relative script path you are 'Dot-Sourcing' 
          the PowerShell function in the script into your PowerShell session which will allow you to 
          run the function 'Get-ModernOfficeApps' from the console.
