###Check Office Update Network Load

Determines the size of the update and quality of delta compression for an office update.

Uses Office Deployment Tool to download and install a specified starting version of Office. 
Then captures the current received bytes on the NetAdapter, before starting an update to 
the specified end version. Records the end received bytes to determine the total size of the 
download. Then zips the apply folder within the office updates folder to determine what the 
max download size would be without delta compression. Comparing these two values provides the 
delta compression value.

[README](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/wiki/README_Check-OfficeUpdateNetworkLoad)

1. Open a PowerShell console.

		From the Run dialog type PowerShell 

2. Change directory to the location where the PowerShell Script is saved.

		Example: cd C:\PowerShellScripts

3. Run the Check-OfficeUpdateNetworkLoad function and specify a GPO Name and the versions of Office to block. 

		Check-OfficeUpdateNetworkLoad -VersionStart 15.0.4623.1003 -VersionEnd 15.0.4631.1002