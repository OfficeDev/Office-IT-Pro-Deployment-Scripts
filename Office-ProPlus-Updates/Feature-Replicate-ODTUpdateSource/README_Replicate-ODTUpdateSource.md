#ODT Server Replication

This script will provide IT Pros a way to manage ODT replication between a source server and a remote destination.

###Pre-requisites

1. A shared folder that will host the C2R builds. This will be referred to as the source.
2. Remote shared folders that will replicate the source.
3. A configuration.xml file used to poll the CDN for C2R builds.
4. The setup.exe used to download the C2R builds from the CDN.

###Links

Overview on using the ODT - https://technet.microsoft.com/en-us/library/jj219422.aspx
    
Download the ODT - http://www.microsoft.com/en-us/download/details.aspx?id=36778

Reference for Click-to-Run configuration.xml files - https://technet.microsoft.com/en-us/library/JJ219426.aspx

Reference for creating scheduled tasks - https://msdn.microsoft.com/en-us/library/windows/desktop/bb736357(v=vs.85).aspx

###Network Share

When deciding the location of the network share you should consider the locations from which the client workstations are accessing the share. There are several options to ensure that workstations are installing Office over their local network.

  1. Multiple GPOs - You could create a separate Group Policy and a Network Share for each network site. For each network site you could create local share for the Office installation file and create a Group Policy that is applied to the workstations in that site to point to that local share. This will provide a solution that provides a local copy of the Office installation files for each site. The limitation to this solution is depending on how the Group Policy is assigned to the workstations this may not ensure that the computer is using a local share to install Office. This could happen if a laptop user is a different location from where their Group Policy is assigned. Another issue with this solution is having to maintain multiple Group Policies.

  2. DFS Shares - If the netowrk share that is used is a Distributed File System (DFS) share you can leverage the replication capabilities of DFS to ensure that each network site has copy of the Office installation files. Also by using a DFS share path you can ensure that the workstations

  3. Netlogon Share - By using the netlogon share on the Active Directory Domain Controllers to store the Office installation files you can ensure the workstations are always using the closest Domain Controller to install Office. Since this solution uses Active Directory replication to copy the Office installation files to every Domain Controller in the Domain you must ensure that every Domain Controller has enough free space, on the volume where the SYSVOL share is located, to store the Office installation files.

###Setup

Copy the files below in to the folder from where the script will be ran.

    Office2013setup.exe
    configuration.xml
    
Use the link provided above for configuring the xml file.

###Functions

####Download-ODTOfficeFiles

#####Parameters

OfficeVersion - The version of Office used for the ODT

XmlConfigPath - Path to the Configuration xml file located on a shared folder

TaskName - The name of the task created on the source computer

#####Example 1

    1. Open PowerShell as an administrator.

          From the Run dialog type PowerShell, right click it and choose Run as Administrator
        
    2. Change the directory to the location where the PowerShell Script is saved. 
  
          Example: cd C:\PowerShellScripts

    3. Dot-Source the script to gain access to the functions inside.

          Type: . .\Manage-ODTReplication.ps1

            By including the additional period before the relative script path you are 'Dot-Sourcing' 
            the PowerShell function in the script into your PowerShell session which will allow you to 
            run the inner functions from the console.
   
    4. Download the latest C2R build with a specified Configuration xml file
  
          Download-ODTOfficeFiles -OfficeVersion 2013 -XmlConfigPath "\\Server1\ODT Replication"
          
#####Example 2

    1. See the first three steps in Example 1.
  
    2. Create a task on the source that will poll the CDN daily and download the latest C2R build.
  
          Download-ODTOfficeFiles -OfficeVersion 2013 -XmlConfigPath "C:\ODT Replication" -TaskName "ODT CDN Poll" -ScheduledTime 03:00

####Replicate-ODTOfficeFiles

#####Parameters

Source - The source folder hosting the C2R builds.

ODTShareNameLogFile - The name of the csv file containing a list of shared folders.

#####Example
  
    1. Open PowerShell as an administrator.

          From the Run dialog type PowerShell, right click it and choose Run as Administrator
        
    2. Change the directory to the location where the PowerShell Script is saved. 
  
          Example: cd C:\PowerShellScripts

    3. Dot-Source the script to gain access to the functions inside.

          Type: . .\Manage-ODTReplication.ps1
          
    4. Compare the remote share to the source folder. If the source folder has updated files or folders the remote share will replicate the source.
  
        Replicate-ODTOfficeFiles -Source "\\Server1\ODT Replication" -ODTShareNameLogFile "\\Server1\ODT Replication\ODTRemoteShares.csv"

####Schedule-ODTRemoteShareReplicationTask

#####Parameters

ComputerName - LIst of computers to create the shceduled task on.

Source - The source share hosting the C2R builds.

TaskName - The name of the scheduled task.

Schedule - A trigger for the script to run Monthly. "MONTHLY" will autopopulate.

Modifier - The value that refines the scheduled frequency. The list of available
modifiers are FIRST,SECOND,THIRD,FOURTH,LAST.

Days - Provide the day of week for the task to run on. The list of available
days are MON,TUE,WED,THU,FRI,SAT,SUN.

StartTime - The time of day the task will run. The hour format is 24-hour (HH:mm)
If no StartTime is given the time will default to the time the task is created.

#####Example

    1. Open PowerShell as an administrator.

          From the Run dialog type PowerShell, right click it and choose Run as Administrator
        
    2. Change the directory to the location where the PowerShell Script is saved. 
  
          Example: cd C:\PowerShellScripts

    3. Dot-Source the script to gain access to the functions inside.

          Type: . .\Manage-ODTReplication.ps1
          
    4. Create a scheduled task on a remote computer.
    
          Schedule-ODTRemoteShareReplicationTask -ComputerName Computer1,Computer2 -Source "\\Server1\ODT Replication" -TaskName "ODT Replication" -Schedule MONTHLY -Modifier SECOND -Days WED -StartTime 03:00 

####Add-ODTRemoteUpdateSource

#####Parameters

ODTShareNameLogFile - The name of the csv file containing a list of shared folders.

RemoteShares - A list of remote shares to remove from the csv.

#####Example
    1. Open PowerShell as an administrator.

          From the Run dialog type PowerShell, right click it and choose Run as Administrator
        
    2. Change the directory to the location where the PowerShell Script is saved. 
  
          Example: cd C:\PowerShellScripts

    3. Dot-Source the script to gain access to the functions inside.

          Type: . .\Manage-ODTReplication.ps1
          
    4. Add a list of remote shares to be recorded in  csv file.
    
          Add-ODTRemoteUpdateSource -RemoteShare "\\Computer3\ODT Replication","\\Computer4\ODT Replication" -ODTShareNameLogFile "\\Server1\ODT Replication\ODTRemoteShares.csv"
    
####Remove-ODTRemoteUpdateSource

#####Parameters

ODTShareNameLogFile - The name of the csv file containing a list of shared folders.

RemoteShares - A list of remote shares to remove from the csv.

#####Example

    1. Open PowerShell as an administrator.

          From the Run dialog type PowerShell, right click it and choose Run as Administrator
        
    2. Change the directory to the location where the PowerShell Script is saved. 
  
          Example: cd C:\PowerShellScripts

    3. Dot-Source the script to gain access to the functions inside.

          Type: . .\Manage-ODTReplication.ps1
          
    4. Remove a remote share from the csv file containing the list of available shares to replicate to.

          Remove-ODTRemoteUpdateSource -ODTShareNameLogFile "\\Server1\ODT Replication\ODTRemoteShares.csv" -RemoteShares "\\Computer1\ODT Replication","\\Computer2\ODT Replication"
          
####List-ODTRemoteUpdateSource

#####Parameter

ODTRemoteUpdateSource - The name of the csv file containing a list of shared folders.

#####Example

    1. Open PowerShell as an administrator.

          From the Run dialog type PowerShell, right click it and choose Run as Administrator
        
    2. Change the directory to the location where the PowerShell Script is saved. 
  
          Example: cd C:\PowerShellScripts

    3. Dot-Source the script to gain access to the functions inside.

          Type: . .\Manage-ODTReplication.ps1 
          
    4. List the available shares to replicate to.

         List-ODTRemoteUpdateSource -ODTShareNameLogFile
