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
basic execution of the script
powershell.exe -ExecutionPolicy Bypass -NoLogo -NonInteraction -NoProfile -WindowStyle Hidden -File .\Create-OfficeAddingRecords.ps1 -sqlServer 'cm01\' -dbName 'CM_S01' -user 'DOMAIN\username' -password 'password'
if you are running the script as a package in SCCM  

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
        $addins = Get-OfficeAddins

        $workstationTable = "Workstations"
        $addinTable = "OfficeAddins"
        $workstationAddinsTable = "WorkstationAddinJunction"

        $connectionString ="Server="+$sqlServer+";Database="+$dbName+";Integrated Security=SSPI;"
        $connection = New-Object System.Data.SqlClient.SqlConnection
        $connection.ConnectionString = $connectionString
        $connection.Open()

        $command = New-Object System.Data.SqlClient.SqlCommand
        $command.Connection = $connection

        InsertWorkStation -workstationName $env:COMPUTERNAME -command $command
        if($addins)
        {
             foreach($addin in $addins){
            
            }
        }

    }
    End{
        $connection.Close()
    }
}

function InsertWorkStation{
    Param(
        [Parameter(Mandatory=$True)]
        [String]$workstationName = "",
        [Parameter(Mandatory=$True)]
        [System.Data.SqlClient.SqlCommand]$command
    )
    $Id = [guid]::NewGuid()
    $command.CommandText = "INSERT INTO Worksations VALUES('"+$Id+"','"+$workstationName+"')"
    $command.EndExecuteNonQuery()
}

function WorkstationExists{
        Param(
        [Parameter(Mandatory=$True)]
        [String]$workstationName = "",
        [Parameter(Mandatory=$True)]
        [System.Data.SqlClient.SqlCommand]$command
        )

        $command.CommandText = "SELECT * FROM Workstations WHERE Name = '"+$workstationName+"'"
        $results = $command.ExecuteNonQuery() 

        if($results){
            return $true
        }
           
        return $false 
    }

function TableExists{
        Param(
        [Parameter(Mandatory=$True)]
        [String]$tableName = "",
        [Parameter(Mandatory=$True)]
        [System.Data.SqlClient.SqlCommand]$command
        )

        try{
            $command.Text = "SELECT * FROM "+$tableName 
            $result = $command.ExecuteNonQuery();
            if($result){
                return $true 
            }
            return $false
            
        }
        catch{
            return $false
        }
    }
