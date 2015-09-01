#**SCCM Office Update Trigger**

This application is designed to be run remotely on a users Office 365 Pro Plus PC, via a System Center Configuration Manager  (SSCM) package containing the source media for an Office 365 Pro Plus build. This ensures that the PC updates Office 365 Pro Plus from that SCCM package on the DP closest to it.

When used with SCCM the script is deployed using an SCCM package.  In order for this executable to properly work the package must be configured with the 'Package share settings' set to 'Copy the content in this package to a package share on distribution points'.  Office Click-To-Run updates cannot be deployed directly with SCCM but instead this script is leveraging the SCCM Distrbution Points (DP) has update sources via a network share.  With SCCM 2012 if you do not configure the setting above then SCCM will use the single instance store and the Office Update files will not be directly available over the network.  When using the traditional package share those files will be available over the network.  In the solution prodided with the [Setup-SCCMOfficeUpdates](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/blob/master/Setup-SCCMOfficeUpdates/README_Setup-SCCMOfficeUpdates.md) in this repository it configures the package to run the executable from the UNC path of the Distribution Point.  When the executable is run it will detect this UNC path and set that in the registry as the Office Update path.  The [Setup-SCCMOfficeUpdates](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/blob/master/Setup-SCCMOfficeUpdates/README_Setup-SCCMOfficeUpdates.md) script will also configure the package to run on the client at a interval so it tries to ensure that mobile users are using their closest DP.

###**Pre-requisites**

1. This script is designed to work with System Center Configuration Manager (SCCM). There is a script in this GitHub repository that automates the process of implementing this executable with SCCM.  For more information on this go to the following link [Setup-SCCMOfficeUpdates](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/blob/master/Setup-SCCMOfficeUpdates/README_Setup-SCCMOfficeUpdates.md).




	

	

