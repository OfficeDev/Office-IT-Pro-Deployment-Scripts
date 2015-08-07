#New GPO Office Installation

This script will configure an existing Active Directory Group Policy to silently install Office 2013 Click-To-Run on computer startup.

###Pre-requisites

1. Active Directory
2. A shared network folder for Office Installation Files.
3. An existing Group Policy Object that is assigned to the target computer you want to install Office 2013 Click-To-Run

###Setup

Copy the files below in to the folder from where the script will be ran.

        Configure-GPOOfficeInstallation.ps1
        Configuration_Download.xml
        Configuration_InstallLocally.xml
        Configuration_template.xml
        InstallOffice2016.ps1
        SetupOffice2013.exe 


###Example

1. Open PowerShell as an administrator.

          From the Run dialog type PowerShell, right click it and choose Run as Administrator

2. Change the directory to the location where the PowerShell Script is saved.
          Example: cd C:\PowerShellScripts
      
3. Dot-Source the script to gain access to the functions inside.

           Type: . .\Configure-GPOOfficeInstallation.ps1

           By including the additional period before the relative script path you are 'Dot-Sourcing' 
           the PowerShell function in the script into your PowerShell session which will allow you to 
           run the inner functions from the console.

4. Run the Download-GPOOfficeInstallation cmdlet and specify the paramaters, -UncPath

          Download-GPOOfficeInstallation -UncPath "\\Pathname\Sharename"
      
   Office will download the Office install files to the specified folder share 
   and will copy the Configuration_Download.xml, 
   Configuration_InstallLocally.xml, and the SetupOffice2013.exe files. 

5. Run the "SetUpOfficeInstallationGpo.ps1" script and specify the paramaters, $UncPath and $GpoName.

          Configure-GPOOfficeInstallation -UncPath "\\Pathname\Sharename" -GpoName "GroupPolicyName"

6. Refresh the Group Policy on a client computer:

          From the Start screen type command and Press Enter
          Type "gpupdate /force" and press Enter.

7. Restart the client computer.

          When the client computer starts the script will launch in the background. 
          You can verify if the script is running by opening 
          Task Manager and look for the Click To Run process.







