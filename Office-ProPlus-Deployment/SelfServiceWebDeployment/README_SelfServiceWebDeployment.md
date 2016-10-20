#Office 365 ProPlus Self Service Deployment (SSDS)
##Deploying to IIS

###Install Web Deploy
1. You must have Windows Server 2012 or greater with **Internet Information Services (IIS)** installed with ASP.NET 4.5 and .NET Extensibility 4.5 installed in IIS.
2. If you currently don’t have the Microsoft Web Platform Installer installed then navigate to http://www.microsoft.com/web/downloads/platform.aspx 
2. Click **Download** and run the installer.
3. Click on the **Products** tab in the Web Platform Installer.
4. For Server 2012 R2 and Greater in the search box type **Web Deploy 3.6 without bundled SQL Support**¸ then click the **Add** option.
5. For Server 2012 in the search box type **Web Deploy 3.5 without bundled SQL Support**¸ then click the **Add** option.
6. Click **Install** at the bottom of the window.
7. Click **I Accept** in the new window.  Wait for the program to be downloaded and installed. 
8. Click the **Finish** button in the new window that displays after your program has been installed.

###Configure IIS
1. Go to your **Start Menu** and type **IIS**, then select **Internet Information Services (IIS) Manager**.
2. Download the deployment package for the website from [Self Service Site Deployment Package](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/raw/master/Office-ProPlus-Deployment/SelfServiceWebDeployment/OfficeProPlusSelfServiceSite.zip)
3. If you want to create a new website to use then right click **Sites** and the click **Add Website**.  If you want to use an existing website then skip to step 8.
4. In the new Window add the name of the new site into **Site Name**.
5. In the **Physical path:** enter an existing local folder path where you want the files to be located. (ex: C:\SelfService)
6. If an existing website is already using port 80 for http or 443 for https then you cannot reuse those ports unless you are using a different IP Address or hostname.  If you want to use ports already 
   in use then consider adding the site to and existing Website.
7. Click **OK**
8. Right click on the website you want to use and the click **Deploy – Import Application**
9. Located the package that you downloaded in step 2 and click **Next**
10. Click **Next**
11. If there is existing content in the website then you should change or use the default Application Path (ie: Office365ProPlus).  If you want the site to be at the root of the website then clear the Application Path (Note: This should only be done for a newly created website).
13. Click **Next**
14. Click **Next**
15. Click **Finish**
16. If you accepted the defaults to install the application navigate to http://servername/Office365ProPlus/SelfService.  If the application was installed in the root of the website navigate to http://hostname/SelfService.

##Configure Windows Firewall
If you are not using a standard port for the website you may have to make changes to the Windows Firewall in order to allow remote computers to access the application.

1. Go to the **Start Menu**.
2. Type **Firewall Advanced** and then select the **Windows Firewall With Advanced Security** option.
3. In the panel on the left hand side of the new window select the **Inbound Rules** option.
4. Search for and double click on the **World Wide Web Services** rule.
5. Select the **Enable** option in the new window then press the **OK** button.
6. In the right hand side of the window select the **New Rule** option.
7. In the new window select the **Port** option then click **Next**.
8. In the **Specific local ports** field enter the port used in step 7 of **Configuring IIS** (ex: port 81) and click **Next**.
9. In the next page make sure that the **Allow the connection** option is selected then click **Next**.
10. Click **Next** in the following page.
11. Enter a name for your rule in the **Name** field then click the **Finish** button.

#Site Configuration
##Configuration XML
The file **SelfServiceConfig.xml** is located at the root of the site and allows for the customization of the SSDS.  The customizable areas of the site are as follows; the company logo in the site’s banner, the 
company’s name in the site’s banner, the questions and answers on the help page, and the builds offered. Each "Build" represents an ODT Configuration XML file which corresponds to a deployment item on the site. 

####Example SelfServiceConfig.xml
![alt text](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/blob/Development/Office-ProPlus-Deployment/SelfServiceWebDeployment/images/ExampleSelfServiceConfigXml.png "Example self service config.xml")

###Company Name
To configure the company simply set **Name** attribute of the **Company** element in the **SelfServiceConfig.xml** file.

###Add a Build
To add a build to the Package Selection page add a **Build** element as a child of the **Builds** element with a **Languages**, **Filters**, **Location**, **Type**, and **ID** attributes. The ID attribute must match the file name exactly of the XML configuration file located in the **“XmlFiles”** directory which is in the root of the website. A user that selects the build will install Office 365 ProPLus with the configuration contained in that Xml File plus the languages that they select.

####Languages Attribute
The Languages attribute should be populated with a comma separated list of the language packs that are available for selection as either the primary language or additional lanuages.

####Possible Language Packs
There is a set list of valid language packs that are available for use.  That list of languages includes:

|                               |                              |                                |         
|-------------------------------|------------------------------|--------------------------------|
| English (en-us)               | Greek (el-gr)                | Polish (pl-pl)                 |
| Arabic (ar-sa)                | Hebrew (he-il)               | Portuguese (Brazil) (pt-br)    |
| Bulgarian (bg-bg)             | Hindi (hi-in)                | Portuguese (Portugal) (pt-pt)  |
| Chinese (Simplified) (zh-cn)  | Hungarian (hu-hu)            | Romanian (ro-ro)               |
| Chinese (zh-tw)               | Indonesian (id-id)           | Russian (ru-ru)                |
| Croatian (hr-hr)              | Italian (it-it)              | Serbian (Latin) (sr-latn-rs)   |
| Czech (cs-cz)                 | Japanese (ja-jp)             | Slovak (sk-sk)                 |
| Croatian (hr-hr)              | Kazakh (kk-kh)               | Slovenian (sl-si)              |
| Danish (da-dk)                | Korean (ko-kr)               | Spanish (es-es)                |
| Estonian (et-ee)              | Latvian (lv-lv)              | Swedish (sv-se)                |
| Finnish (fi-fi)               | Lithuanian (lt-lt)           | Thai (th-th)                   |
| French (fr-fr)                | Malay (ms-my)                | Turkish (tr-tr)                |
| German (de-de)                | Norwegian (nb-no)            | Ukrainian (uk-ua)              |          

####Filters Attribute
The Filters attribute is used to help further differentiate builds from one another.  It can be populated by a comma separated list of arbitrary values.  These values are displayed as Tags 
in the tooltip of builds when viewed in the panel format or as values in the Tags column when viewed in the table/list format on the Package Selection page. 
For example, the first Build in the example **SelfServiceConfig.xml** uses the following Filters: Level III and FullTime.

####Location Attribute
The Location attribute is yet another  attribute used to differintiate each build from one another.  The Location attribute can be populated by any arbitrary value thate the administrator wishes to set.  
These values are displayed as the Location section of each build when the builds are viewed in the panel format or as a value in the Location column when viewed in the table/list format on the Package Selection page. 
For example, the first Build in the example **SelfServiceConfig.xml** uses the following Location: Florida.

####Type Attribute
The Type attribute is used as the title of each build when displayed on the Package Selection page.  The Type attribute is displayed above the Location attribute when the Package Selection page is viewed 
in panel view and as the value of the Name column when the Package Selection page is viewed in the table/list format. 
For example, the first Build in the example **SelfServiceConfig.xml** uses the following Type: IT Pro.

####ID Attribute
The ID attribute is used to correlate each build on the Package Selection page with a partially complete XML configuration file located in the **“XmlFiles”** directory which is located at the root of the website. 
For example, the first Build in the example **SelfServiceConfig.xml** uses the following ID: ExecutivesNewYork.  If you look in the **XmlFiles** folder there should be a file name ExecutivesNewYork.xml which contains ODT Configuration xml.

###Help Page Content
To add content to the Help Page add an **Item** element as a child of the **Help** element with a **Question** and **Answer** element as children.  Add the possible question as the contents of the **Question** 
element and then add the answer to this question as the contents of the **Answer** element.

#Build Configuration
##Base Build Files
Each build displayed on the Package Selection page must have partially completed XML configuration file with a file name corresponding to its ID attribute in the **SelfServiceConfig.xml** 
located in the **“XmlFiles”** directory.  These base configuration files are modified according to the languages selected by the user and then copied to the 
**“Content\Generated_Files”** directory.  The base configuration file can be generated using the tool found [here](http://officedev.github.io/Office-IT-Pro-Deployment-Scripts/XmlEditor.html).

####Example ExecutivesNewYork.xml File  
![alt text](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/blob/Development/Office-ProPlus-Deployment/SelfServiceWebDeployment/images/ExampleBaseBuild0XmlFile.png "Example ExecutivesNewYork.xml")

###Generated Build Files
Generated build files are the combination of the languages selected when the user selects their primary/addiitonal languages and the base build file associated with the selected build.  
The generated build file is used by the ClickOnce Installer to install the correct build and requested language packs. 

#Basic Site Usage
##Package Selection
When the user first loads the SSDS (using Internet Explorer or Microsoft Edge) they will be brought to the package selection page (screenshot below).  This page displays all of the 
packages that are available for installation as well as giving the user the ability to search and filter these builds by certain criteria.
![alt text](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/blob/Development/Office-ProPlus-Deployment/SelfServiceWebDeployment/images/PackageSelection.png "Package selection")

###Packages
Packages are pre-defined builds that are created by the site’s administrator.  They allow for the tailoring of the builds for specific users.  The builds are differentiated by three 
different fields, the build name, the build location, and build tags.  All three of these fields are customizable and populated using the **SelfServiceConfig.xml** file (this file will be covered later in the documentation). 

####Tile View
![alt text](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/blob/Development/Office-ProPlus-Deployment/SelfServiceWebDeployment/images/TileView.png "Tile view")

####List View
![alt text](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/blob/Development/Office-ProPlus-Deployment/SelfServiceWebDeployment/images/ListView.png "List view")

####Package Filtering
Users are able to filter the displayed builds by using the Live Searchbox as well as the Location Dropdown.  The Live Searchbox can be used to filter by any of the three different fields 
contained by each build.  The Location Dropdown can only filter by the Location field. 
![alt text](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/blob/Development/Office-ProPlus-Deployment/SelfServiceWebDeployment/images/PackageFiltering.png "Package filtering")

####View Toggling
Users are able to toggle the builds that are displayed.  They can either be viewed in a tile format by pressing the Tile View button or in a list/table format by pressing the List View button.
![alt text](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/blob/Development/Office-ProPlus-Deployment/SelfServiceWebDeployment/images/ViewToggling.png "View toggling")

####Selecting A Package For Installation
To select a build for installation simply click the “Install” text associated with that tile/list item.  Once the “Install” text has been selected the user will be taken to a page that 
requires them to select a primary language for their installation.
![alt text](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/blob/Development/Office-ProPlus-Deployment/SelfServiceWebDeployment/images/SelectingAPackageForInstallation1.png "Select a package for installation")
![alt text](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/blob/Development/Office-ProPlus-Deployment/SelfServiceWebDeployment/images/SelectingAPackageForInstallation2.png "Select a package for installation")

##Language Selection
###Primary Language Selection
The language selected on this page will be the language used in the installer as well as the default language used by the programs included in this build.  A primary language must be 
selected before being able to proceed to the next page. 
![alt text](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/blob/Development/Office-ProPlus-Deployment/SelfServiceWebDeployment/images/PrimaryLanguageSelection.png "Primary language selection")

###Additional Language Selection
Once a primary language has been selected the user will be taken to the Additional Languages page.  Here the user can select any other language packs that they wish to install next 
to primary language.  Additional languages are optional and may be skipped by the user.  Note that the additional language options as well as the primary language options must be 
associated with the selected build in the **SelfServiceConfig.xml** file.  
![alt text](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/blob/Development/Office-ProPlus-Deployment/SelfServiceWebDeployment/images/AdditionalLanguageSelection.png "Additional language selection")

#The Installation Process
##Beginning the Installation Process
After a user has selected any additional languages they are brought to a confirmation page, listing out the package’s information along with the chosen languages.  When the user clicks the Install button (A) SSDS will generate a configuration file using a base xml file associated with the selected build along with the languages that the user selected, then a ClickOnce installer will be downloaded and the installation process will begin.
![alt text](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/blob/Development/Office-ProPlus-Deployment/SelfServiceWebDeployment/images/BeginningTheInstallationProcess.PNG "Beginning the installation process")

##ClickOnce Installer
Once the user has selected the Install button on the previous page the ClickOnce installer will be downloaded.  This requires that pop-ups are allowed for the SSDS.  Once downloaded, the user will need to click the Install button.  the ClickOnce installer to download the Office 365 ProPlus 2016 installer as well as the custom generated installation configuration file.   
####ClickOnce Download
![alt text](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/blob/Development/Office-ProPlus-Deployment/SelfServiceWebDeployment/images/ClickOnceDownload.PNG "ClickOnce download")

##Using the ClickOnce Installer
After the ClickOnce Installer has been downloaded the user will be prompted to begin the installation procces.  They will first need to click the Install button.  After the user presses the Install button they may be notified that the ClickOnce installer is unrecognized (this is because it is unsigned).  They will need to select the More Info option and then press the Run anyway option to allow for the installer to run.  The Office Installer will then be launched.  

####ClickOnce Installer
![alt text](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/blob/Development/Office-ProPlus-Deployment/SelfServiceWebDeployment/images/ClickOnceInstaller.PNG "ClickOnce installer")

####Office Install
![alt text](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/blob/Development/Office-ProPlus-Deployment/SelfServiceWebDeployment/images/OfficeInstall.PNG "Office install")

##Possible Issues
###Missing Configuration File
If a user attempts to install a package that does not have a base configuration xml file located on the server, the following dialog will be shown when attempting to download the installer.  If this occurs the site’s administrator will need to add a configuration xml file with the correct name, in this case “build1.xml”, to the “Content/XML_Build_Files/Base_Files/” directory of the SSDS.

![alt text](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/blob/Development/Office-ProPlus-Deployment/SelfServiceWebDeployment/images/MissingConfigurationFile.PNG "Missing configuration file")
