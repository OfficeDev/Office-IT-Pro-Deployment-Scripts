#Lock Office Applications

This PowerShell function will prevent a client in the domain from opening a specified version of Office by creating a group policy and WMI filter.

###Pre-requisites
1. Active Directory
2. Server Manager
3. Copy the New-OfficeAppLockGPO and Manage-OfficeWmiFilters functions to a local folder.

###Links
Group Policy Management Console - https://technet.microsoft.com/en-us/library/cc753298.aspx
WMI filtering - https://technet.microsoft.com/en-us/library/cc779036(v=ws.10).aspx

###**Examples**

1. Open a PowerShell console.

		From the Run dialog type PowerShell 

2. Change directory to the location where the PowerShell Script is saved.

		Example: cd C:\PowerShellScripts

3. Dot-Source the New-OfficeAppLockGPO and Manage-OfficeWmiFilters functions into your current session.

		Type . .\New-OfficeAppLockGPO.ps1
		Type . .\Manage-OfficeWmiFilters.ps1
		
		By including the additional period before the relative script path you are 'Dot-Sourcing' 
		the PowerShell function in the script into your PowerShell session which will allow you to 
		run the function from the console.

4. Run the New-OfficeAppLockGPO function and specify a GPO Name and the versions of Office to block.
              
        New-OfficeAppLockGPO -GpoName "Lock Office 2010,2013" -OfficeVersion Office2010,Office2013

        Note: After dot sourcing the New-OfficeAppLockGPO the available versions of Office to block 
        will auto-populate after typing -OfficeVersion. The available versions are Office2003, Office2007, 
        Office2010, and Office2013.
        
5. Link the new GPO to an appropriate OU.

        1. Open the Group Policy Management console
        2. Refresh the WMI Filters by right clicking and choose Refresh.
        3. Refresh the Group Policy Objects by right clicking and choose Refresh.
        4. Right click an appropriate OU and choose "Link an Existing GPO..."
        5. Highlight the new GPO and click OK.
