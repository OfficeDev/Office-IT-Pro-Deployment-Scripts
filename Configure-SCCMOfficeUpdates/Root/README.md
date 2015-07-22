##**Update Office 2013 or Office 2016 using SCCM**

Automates the process of updating an existing Office 2013 or Office 2016 installation controlled through SCCM

###**Pre-Requisites:**

Before running this script, the following conditions have to be met

1. .Net Framework 3.5 SP1 must be installed on client machines.
2. You have a functioning SCCM environment set up.
3. Office 15 or Office 16 is already installed on client machines. 
4. Office Auto Updates have been Disabled on the client machines preferably via Group Policy

###**Assumptions:**

1. It is assumed that for this scenario, the client machines will have one of the following OSs - Windows 7, Windows 8, Windows 8.1, Windows 10.
2. The script defaults to use 64 bit version of Office, this can be changed by using the appropriate optional parameter. 

###**Terms:**

1. *Version* - Office Monthly build version number e.g. "15.0.4727.1003" to which you wish to update.
2. *Share* - A UNC path where the office update bits will be stored, this is where the target clients will pull the data to update
3. *SiteId* - The 3 Letter Site ID, used to connect the SCCM PowerShell Session.

###**Files:**

1. **configuration_template.xml** - Base config file template, controls the behaviour office update.

2. **configuration_UpdateSource.xml** - Sample config file used to download bits for the specified Office Build / Version

3. **configuration_UpdateTestGroup.xml** - Sample config file used to update the target test machines to the specified Office Build / Version

4. **SetupOfficeUpdatesSCCM.ps1** - The main script file. Updates the share with the correct configuration files, then downloads the bits, and finally, calls SetupOfficeUpdates.ps1 to create the required SCCM Automation.

5. **SetupOfficeUpdates.ps1** - Creates a Package in SCCM to update Office on target Device Collection, Creates a Program Definition to run the actual binaries from Target Client Machines, 
   Copies package content to a distribution point group, and kicks off the Deployment.

6. **SCO365PPTrigger.exe** - This executable is used to trigger the Office Update correctly on the target client machine. Must be placed in the *Share* location.

7. **setup.exe** - This executable is used to download the bits on the Share as well as update the office installation on the target. Must be copied to the *Share* location.

###**Running the script**

1. Open a Elevated PowerShell Console(see, [Starting Windows PowerShell](https://technet.microsoft.com/en-us/library/hh857343.aspx)):

	```
	From the Run dialog type PowerShell.
	```

2. Change directory to the location where the PowerShell Script is saved.
```
		Example: cd C:\PowerShellScripts
```
   This directory must contain all the *configuration_UpdateSource.xml*, *configuration_UpdateTestGroup.xml* files mentioned above, along with *setup.exe*, and both the *.ps1* files.

3. Run the following in an elevated PowerShell Session
4. Type
```PowerShell
		 . .\SetupOfficeUpdatesSCCM.ps1 -version "*Version*" -path "Share" -siteId "SiteId"
```
4. Monitor the Content Distribution, and the Deployment for status.
