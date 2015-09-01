#Check Office Update Network Load

Determines the size of the update and quality of delta compression for an office update.

Uses Office Deployment Tool to download and install a specified starting version of Office. 
Then captures the current received bytes on the NetAdapter, before starting an update to 
the specified end version. Records the end received bytes to determine the total size of the 
download. Then zips the apply folder within the office updates folder to determine what the 
max download size would be without delta compression. Comparing these two values provides the 
delta compression value.

###Pre-requisites

Recommend using Clean VM

###Examples

1. Set-up a VM and ensure it has an internet connection

2. Open a PowerShell console.

          From the Run dialog type PowerShell and press Enter.
  
3. Change the directory to the location where the PowerShell Script is saved.

          Example: set-location C:\PowerShellScripts
  
4. Run the Script.

          Type: ./Check-OfficeUpdateNetworkLoad -VersionStart 15.0.4623.1003 -VersionEnd 15.0.4631.1002
           

