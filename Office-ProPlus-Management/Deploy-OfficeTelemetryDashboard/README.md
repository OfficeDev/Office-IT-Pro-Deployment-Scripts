### Deploy Office Telemetry Dashboard
This script installs and enables the telemetry agent on computers. 

**IT Pro Scenario:** IT Pros looking for an overview of the number of documents installed on a machine, the number of add ins, etc. Provides an overall health of client computers.
 
[README](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/wiki/README_Deploy-OfficeTelemetry)


1. Open a PowerShell console.

		From the Run dialog type PowerShell 

2. Change directory to the location where the PowerShell Script is saved.

		Example: cd C:\PowerShellScripts

3. Run the script Configure-TelemetryGpo.ps1 to create the GPO	

		Example: ./Create-TelemetryGpo -GpoName "Office Telemetry" -CommonFileShare "Server1" -officeVersion 2013
			A GPO named "Office Telemetry" will be created. Registry keys will be created to enable telemetry agent logging, uploading, and the commonfileshare path set to \\Server1\TDShared. l


