\#New GPO Office Installation

This script will configure an existing Active Directory Group Policy to silently install Office 2013 Click-To-Run on computer startup.

###Pre-requisites

1. Active Directory
2. A shared network folder for Office Installation Files.
3. An existing Group Policy Object that is assigned to the target computer you want to install Office 2013 Click-To-Run

###Setup

Copy the files below in to the folder from where the script will be ran.

        configuration_Download.xml
        Configuration_InstallLocally.xml
        configuration_template.xml
        DownloadOfficeInstallationToNetworkShare.ps1
        InstallOffice2016.ps1
        SetUpOfficeInstallationGpo.ps1
        SetupOffice2013.exe 


###Example

1. Open PowerShell as an administrator.

          From the Run dialog type PowerShell, right click it and choose Run as Administrator

2. Change the directory to the location where the PowerShell Script is saved.
          Example: cd C:\PowerShellScripts
      
3. Run the "DownloadOfficeInstallationToNetworkShare.ps1" script and specify the paramaters, $UncPath and $Bitness.

          . .\DownloadOfficeInstallationToNetworkShare -UncPath "\\Pathname\Sharename" -Bitness 32
      
   Office will download per the bit specified to the folder share 
   and will copy the Configuration_Download.xml, 
   Configuration_InstallLocally.xml, and the setup.exe files. 
   The xml files will reflect the bit specified next to OfficeClientEdition.

4. Run the "SetUpOfficeInstallationGpo.ps1" script and specify the paramaters, $UncPath and $GpoName.

          . .\SetUpOfficeInstallationGpo -UncPath "\\Pathname\Sharename" -GpoName "MyGpo"
      
          The InstallOffice2016.ps1 script will be copied to the GUID 
          located at %systemroot%\SYSVOL\sysvol\domain\Policies.

5. Verify the Startup script in the Group Policy Object:

          1. From within Group Policy Management right click the GPO and choose Edit.
          2. Under Computer Configuration click the Policies drop down.
          3. Expand Windows Settings and click on Scripts.
          4. In the viewer window double click Startup.
          5. Click the PowerShell Scripts tab and verify the PS script and parameters are available.
          6. Click OK to close the Startup Properties window.

6. Refresh the Group Policy on a client computer:

          From the Start screen type command and Press Enter
          Type "gpupdate /force" and press Enter.

7. Restart the client computer.

          When the client computer starts the script will launch in the background. 
          You can verify if the script is running by opening 
          Task Manager and look for the Click To Run process.







