#New GPO Office Installation

This script will configure an existing Active Directory Group Policy to silently install Office 2013 Click-To-Run on computer startup.

For more information on deploying Office via Group Policy go to https://technet.microsoft.com/en-us/library/Ff602181.aspx

###Pre-requisites

1. Active Directory
2. A shared network folder for Office installation files
3. An existing Group Policy Object that is assigned to the target computers you want to install Office 2013 Click-To-Run

###Network Share

When deciding the location of the network share you should consider the locations from which the client workstations are accessing the share.  There are several options to ensure that workstations are installing Office over their local network.

1. **Multiple GPOs** - You could create a separate Group Policy and a Network Share for each network site. For each network site you could create local share for the Office installation file and create a Group Policy that is applied to the workstations in that site to point to that local share.  This will provide a solution that provides a local copy of the Office installation files for each site.  The limitation to this solution is depending on how the Group Policy is assigned to the workstations this may not ensure that the computer is using a local share to install Office.  This could happen if a laptop user is a different location from where their Group Policy is assigned.  Another issue with this solution is having to maintain multiple Group Policies.

2. **DFS Shares** - If the netowrk share that is used is a Distributed File System (DFS) share you can leverage the replication capabilities of DFS to ensure that each network site has copy of the Office installation files.  Also by using a DFS share path you can ensure that the workstations 

3. **Netlogon Share** - By using the netlogon share on the Active Directory Domain Controllers to store the Office installation files you can ensure the workstations are always using the closest Domain Controller to install Office.  Since this solution uses Active Directory replication to copy the Office installation files to every Domain Controller in the Domain you must ensure that every Domain Controller has enough free space, on the volume where the SYSVOL share is located, to store the Office installation files.

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







