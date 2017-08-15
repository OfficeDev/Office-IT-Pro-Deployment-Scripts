function Create-OfficeAddinRecords{


<#
.SYNOPSIS
This script is deployed as a package to clients via SCCM to record the office addins
installed.  

.DESCRIPTION
The scrip is used to 
create records in a sql database pertaining to the addins installed on the client machine
to be used as data for reports. 

.PARAMETER sqlServer
The name of the SCCM server

.PARAMETER dbName
The database name that the data will be recorded in 

.PARAMETER user
User name used to access the sql server

.PARAMETER password
Password of the account used to access the sql server 

.EXAMPLE
Generate-ODTLanguagePackXML -TargetFilePath $env:temp\LanguagePacks.xml -Languages de-de,es-es,fr-fr -OfficeClientEdition 64
A new xml file will be created in the temp directory called LanguagePacks.xml which will be used to install the 64-bit
editions of German, Spanish, and French language packs.

.EXAMPLE
Create-OfficeAddinRecords -sqlServer cm01\ -dbName CM_S01 -user DOMAIN\username -password password 


.NOTES
Date created: 08-15-2017
#>
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$sqlServer = "",
        [Parameter(Mandatory=$True)]
        [String]$dbName = "",
        [Parameter(Mandatory=$True)]
        [String]$user = "",
        [Parameter(Mandatory=$True)]
        [String]$password = "" 
    )
    Begin{
        .\Get-OfficeAddins.ps1
    }
    Process{
    
    }





}