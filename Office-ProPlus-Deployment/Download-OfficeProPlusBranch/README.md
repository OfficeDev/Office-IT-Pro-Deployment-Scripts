#Office ProPlus Branch Downloader

This script will allow you to directly download Office ProPlus 2016 branches to a single location.

For more information on Office ProPlus Branches go to https://technet.microsoft.com/en-us/library/mt455210.aspx

###Branch Management

Office 365 ProPlus (2016) introduced the concept of Branches with allows for greater control over the update process.  Most companies will likely have a mix of Branches that they will use.  Managing updates sources for each branch can be administrative burden.  This script will allow you to easily download the latest build for each branch with a single operation.

###Example

1. Open PowerShell as an administrator.

          From the Run dialog type PowerShell, right click it and choose Run as Administrator

2. Change the directory to the location where the PowerShell Script is saved.
          Example: cd C:\PowerShellScripts
      
3. Dot-Source the script to gain access to the functions inside.

           Type: . .\Download-OfficeProPlusChannels.ps1

           By including the additional period before the relative script path you are 'Dot-Sourcing' 
           the PowerShell function in the script into your PowerShell session which will allow you to 
           run the inner functions from the console.

4. Run the Download-OfficeProPlusBranch cmdlet and specify the paramaters, -TargetDirectory

          Download-OfficeProPlusChannels -TargetDirectory C:\UpdateSource

[![Analytics](https://ga-beacon.appspot.com/UA-70271323-4/README_Download-OfficeProPlusChannels?pixel)](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts)
