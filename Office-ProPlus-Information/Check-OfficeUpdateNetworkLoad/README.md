###Check Office Update Network Load

Determines the size of the update and quality of delta compression for an office update.

Uses Office Deployment Tool to download and install a specified starting version of Office. 
Then captures the current received bytes on the NetAdapter, before starting an update to 
the specified end version. Records the end received bytes to determine the total size of the 
download. Then zips the apply folder within the office updates folder to determine what the 
max download size would be without delta compression. Comparing these two values provides the 
delta compression value.

[README](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/wiki/README_Check-OfficeUpdateNetworkLoad)
