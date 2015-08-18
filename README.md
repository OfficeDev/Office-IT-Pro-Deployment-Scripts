# Office IT Pro Deployment Scripts
This GitHub repository is a collection of useful PowerShell scripts to make deploying Office 2016 and Office 365 ProPlus easier for IT Pros and administrators. 

## Scripts
For more detailed documentation of each script, check the readme file in the corresponding folder

### Edit-OfficeConfigurationFile
Script to modify the Configuration.xml file to configure installation options for Click-to-Run for Office 365 products.

[README](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/blob/master/Edit-OfficeConfigurationFile/README_Edit-OfficeConfigurationFile.md)

### Setup-SCCMOfficeUpdates
Configures System Center Configuration Manager (SCCM) to configure Office Click-To-Run clients to use their closest Distribution Point (DP) for Office Updates.

[README](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/blob/master/Setup-SCCMOfficeUpdates/README_Setup-SCCMOfficeUpdates.md)

### Get-OfficeVersion    
Query a local or remote workstations to find the version of Office that is installed. 

[README](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/blob/master/Get-OfficeVersion/README_Get-OfficeVersion.md)

### Copy-OfficeGPOSettings
Automate the process of moving from an existing version of Office to a newer version while retaining the current set of group policies. 

[README](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/blob/master/Copy-OfficeGPOSettings/README_Copy-OfficeGPOSettings.md)

### Get-OfficeModernApps
Remotely verify the modern apps installed on client machines across a domain.

[README](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/blob/master/Get-ModernOfficeApps/README_Get-ModernOfficeApps.md)

### Configure-GPOOfficeInstallation
This script will configure an existing Active Directory Group Policy to silently install Office 2013 Click-To-Run on computer startup.

[README](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/blob/master/Configure-GPOOfficeInstallation/README_New-GPOOfficeInstallation.md)

### Get-NewOfficeUsers
Two functions to identify licensed Office 365 users and track the dates they were enabled or disabled.

[README](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/blob/master/Get-NewOfficeUsers/README_Get-NewOfficeUsers.md)

## New to PowerShell and Office 365?
Check out [PowerShell for Office 365](https://poweshell.office.com) for advice on getting started, key scenarios and script samples.  

##Questions and comments
If you have any trouble running this sample, please log an issue.
For more general feedback, send an email to o16scripts@microsoft.com.

## How to Contribute to this project
This is high level plan for contributing and the structure that we have in place for pulling changes
<UL>
<LI>There will be 3 main levels of branches: 1 master branch, 1 development branch, feature and bug branches
<LI>Each powershell script will have its own branch and changes will be made at that level
<UL>
<LI>The 3rd level naming conventions will be as follows - Feature-FeatureName or Bug-BugName</UL>
<LI>Pull requests will be made from the feature branches into the development branch and a code review will be completed in the development branch
<LI>Pull requests should only be made from the feature branch after the script is tested and useable
<LI>After the code review is complete a pull request will be made from the development branch into the master branch
<LI>Changes to scripts (new functionality or bug fix) should be done at the thrid level (feature branches) by cloning the development branch using the naming conventions above
<LI>Requests for changes to scripts can be made by submitting an issue and using the appropriate tag
<UL>
<LI>For additional features to an existing script, use the label "enhancement"
<LI>For bugs, use the label "bug"
<LI>All issues will be reviewed and prioritized each day as we work to add new scripts and improve existing ones</UL>
</UL>
