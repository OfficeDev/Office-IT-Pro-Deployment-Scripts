@{

RootModule = 'OfficeProvider.psm1'
ModuleVersion = '1.0.0.1'
GUID = '5ee17516-8006-47e3-bfd9-5d3e4306efc0'
Author = 'Valorem'
CompanyName = 'Valorem'
Description = 'OfficeProvider allows users to install Microsoft Office365 ProPlus from Powershell.'
PowerShellVersion = '3.0'
FunctionsToExport = @()
PrivateData = @{"PackageManagementProviders" = 'OfficeProvider.psm1'

    PSData =@{


        Tags = @("PackageManagement","Provider")

          ReleaseNotes = 'This is a packagemanagement provider that allows users to install Office365 ProPlus.  The provider allows the users to 
          select the bitness, lanaguage packs, and distribution channel of their installation.'

        } # End of PSData
    }
}

