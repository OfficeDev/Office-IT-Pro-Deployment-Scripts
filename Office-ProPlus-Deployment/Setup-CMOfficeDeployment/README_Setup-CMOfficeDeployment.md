### Setup Config Manager Office Deployment

This PowerShell function can setup Config Manager with various Office 365 Client scenarios 

[README](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/wiki/Readme_Setup-CMOfficeDeployment)

The following functions are included with the script.

Download-CMOOfficeChannelFiles
Create-CMOfficePackage
Update-CMOfficePackage
	Create-CMOfficeDeploymentProgram
	Create-CMOfficeChannelChangeProgram
	Create-CMOfficeRollBackProgram
	Create-CMOfficeUpdateProgram
	Create-CMOfficeUpdateAsTaskProgram
Distribute-CMOfficePackage
Deploy-CMOfficeProgram

Scenaro: Install Office

1) Download-CMOOfficeChannelFiles
2) Create-CMOfficePackage
3) Create-CMOfficeDeploymentProgram
4) Distribute-CMOfficePackage
5) Deploy-CMOfficeProgram

Scenario: Channel Change

1) Download-CMOOfficeChannelFiles
2) Create-CMOfficePackage
3) Create-CMOfficeChannelChangeProgram
4) Distribute-CMOfficePackage
5) Deploy-CMOfficeProgram

Scenario: Rollback

For roll back you need to have the version in source to roll back to.

1) Download-CMOOfficeChannelFiles
2) Create-CMOfficePackage
3) Create-CMOfficeRollBackProgram
4) Distribute-CMOfficePackage
5) Deploy-CMOfficeProgram

Scenario: Update Office

1) Download-CMOOfficeChannelFiles
2) Create-CMOfficePackage
3) Create-CMOfficeUpdateProgram
4) Distribute-CMOfficePackage
5) Deploy-CMOfficeProgram

Scenario: Update Office

1) Download-CMOOfficeChannelFiles
2) Create-CMOfficePackage
3) Create-CMOfficeUpdateAsTaskProgram
4) Distribute-CMOfficePackage
5) Deploy-CMOfficeProgram

