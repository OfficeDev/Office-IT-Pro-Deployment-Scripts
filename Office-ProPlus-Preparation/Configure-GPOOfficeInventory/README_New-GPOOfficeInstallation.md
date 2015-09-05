#Configure GPO Office Inventory

This script will configure an existing Group Policy Object (GPO) to schedule a task on workstations to query the version of Office that is installed on the computer and write that information to an attribute on the computer object in Active Directory.

If you don't have System Center Configruration Manager (SCCM) or an equivalent software management system then using this script will provide the capability to inventory what versions of Office are installed in the domain.

With this script the parameter **AttributeToStoreOfficeVersion** will set the attribute on the computer object in Active Driectory that is used to store the Office version.  By default the attribute used is the **info** attribute.  If you want to use an attribute that can be added to the list view in Active Directory User and Computer then two possible attributes would be **telephoneNumber** (Business Phone) and **physicalDeliveryOfficeName** (Office)

In order for this script to work the computer object's **SELF** must have write permissions to the attribute specified.  By default a computer in Active Directory has permissions to write to attributes that are classified as 'Personal Information'.  This functionality is what allows this inventory functionality to work.  The scheduled task that runs on the computer runs under the 'System' context which gives it permissions to write to its own computer account in Active Directory.  If you would like to use an attribute that is not in the 'Personal Information' list then you would have to give 'Self' permissions to write to that Attribute on computer object in Active Directory. A list of possible attributes that you can use are listed below.  The default attribute that is used by this script is Info.  It is an attribute that is unlikely to be already used.  The drawback to using it is that you can not see the value in the computer list view in Active Directory Users and computers.

    -info
    -physicalDeliveryOfficeName
    -assistant
    -facsimileTelephoneNumber
    -InternationalISDNNumber
    -personalTitle
    -otherIpPhone
    -ipPhone
    -primaryInternationalISDNNumber
    -thumbnailPhoto
    -postalCode
    -preferredDeliveryMethod
    -registeredAddress
    -streetAddress
    -telephoneNumber
    -teletexTerminalIdentifier
    -telexNumber
    -primaryTelexNumber

###Pre-requisites

1. Active Directory
3. An existing Group Policy Object that is assigned to the target computers you want to inventory the Office Version

###Setup

Copy the files below in to the folder from where the script will be ran.

        Configure-GPOOfficeInventory.ps1
        Inventory-OfficeVersion.ps1
        ScheduledTasks.xml
        Files.xml

###Example

1. Open PowerShell as an administrator.

          From the Run dialog type PowerShell, right click it and choose Run as Administrator

2. Change the directory to the location where the PowerShell Script is saved.

          Example: cd C:\PowerShellScripts

3. Dot-Source the script to gain access to the functions inside.

           Type: . .\Configure-GPOOfficeInventory

           By including the additional period before the relative script path you are 'Dot-Sourcing' 
           the PowerShell function in the script into your PowerShell session which will allow you to 
           run the inner functions from the console.

4. Run the Configure-GPOOfficeInventory function. This function will configure the Group Policy Object (GPO) specified to inventory the Office version of the workstations that are in the scope of the GPO.

          Configure-GPOOfficeInventory -GpoName WorkstationPolicy

6. Once the computers in scope of the conigured Group Policy Object (GPO) have the policy you should start seeing the Office Versions showing up in Active Directory

7. To refresh the Group Policy on a client computer:

          From the Start screen type command and Press Enter
          Type "gpupdate /force" and press Enter.

8. Exporting computer Office versions from Active Diretory

          Export-GPOOfficeInventory





