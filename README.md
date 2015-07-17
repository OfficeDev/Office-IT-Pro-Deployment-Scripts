# Office IT Pro Deployment Scripts
This repo is a collection of useful PowerShell scripts to make deploying Office 2016 and Office 365 ProPlus easier for IT Pros and administrators. 

## Scripts
For more detailed documentation of each script, check the readme file in the corresponding folder

### Get-OfficeVersion
Query a local or remote workstations to find the version of Office that is installed.

### Get-ModernApps
Remotely verify the modern apps installed on client machines across a domain.

### Copy-OfficePolicies
Automate the process of moving from an existing version of Office to a newer version while retaining the current set of group policies. 

## New to PowerShell and Office 365?
Check out [PowerShell for Office 365](https://poweshell.office.com) for advice on getting started, key scenarios and script samples.  

##Questions and comments
If you have any trouble running this sample, please log an issue.
For more general feedback, send an email to o16scripts@microsoft.com.

## How to Contribute to this project
This is high level plan for contributing and the structure that we have in place for pulling changes.
<UL>
<LI>There will be 3 main levels of branches: 1 master branch, 1 development branch, feature and bug branches
<LI>Each powershell script will have its own branch and changes will be made at that level
<UL>
<LI>The 3rd level naming conventions will be as follows - Feature-FeatureName or Bug-BugName</UL>
<LI>Pull requests will be made from the feature branches into the development branch and a code review will be completed in the development branch
<LI>Pull requests should only be made from the feature branch after the script is tested and useable
<LI>After the code review is complete a pull request will be made from the development branch into the master branch
<LI>Changes to scripts (new functionality or bug fix) should be done at the thrid level (feature branches) by cloning the development branch using the naming conventions above
</UL>
