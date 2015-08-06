Param
(
    [Parameter(Mandatory=$true)]
    [string]$GpoName,
    
    [Parameter()]
    [string]$SqlServerName,

    [Parameter()]
    [string]$OfficeVersion
)


<#

.SYNOPSIS
Create the Telemetry GPO on the Domain Controller

.DESCRIPTION
Creates a group policy that that specifies the 
Telemetry agent file share location and allows
the agent to log and upload.

.PARAMETER GpoName
The name of the GPO to be created.

.PARAMETER SqlServerName
The name of the SQL server hosting the telemetry database.

.PARAMETER OfficeVersion
The version of office used in your environment. If a version
earlier than 2013 is used do not use this parameter.

.EXAMPLE
./Create-TelemetryGpo -GpoName "Office Telemetry" -SqlServerName SQLExpress -officeVersion 2013
A GPO named "Office Telemetry" will be created. Registry keys will be
created to enable telemetry agent logging, uploading, and the commonfileshare 
path set to \\SQLExpress\TDShared. 

.EXAMPLE
./Create-TelemetryGpo -GpoName "Office Telemetry" -SqlServerName SQLExpress
A GPO named "Office Telemetry" will be created.

#>

    Import-Module -Name grouppolicy 
    
    New-GPO -Name $gpoName

    $shareName = "TDShared"
    
    if($OfficeVersion -eq 2013)
    {
        Write-Host "Set the Fileshare name"
        Set-GPRegistryValue -Name "Office Telemetry" -Key "HKCU\Software\Policies\Microsoft\office\15.0\osm" -ValueName CommonFileShare -Type String -Value "\\$SqlServerName\$shareName"

        Write-Host "Enable agent logging"
        Set-GPRegistryValue -Name "Office Telemetry" -Key "HKCU\Software\Policies\Microsoft\office\15.0\osm" -ValueName Enablelogging -Type Dword -Value 1

        Write-Host "Enable agent data upload"
        Set-GPRegistryValue -Name "Office Telemetry" -Key "HKCU\Software\Policies\Microsoft\office\15.0\osm" -ValueName EnableUpload -Type Dword -Value 1
    }
        elseif($OfficeVersion -eq 2016)
        {
            Write-Host "Set the Fileshare name"
            Set-GPRegistryValue -Name "Office Telemetry" -Key "HKCU\Software\Policies\Microsoft\office\16.0\osm" -ValueName CommonFileShare -Type String -Value "\\$SqlServerName\$shareName"

            Write-Host "Enable agent logging"
            Set-GPRegistryValue -Name "Office Telemetry" -Key "HKCU\Software\Policies\Microsoft\office\16.0\osm" -ValueName Enablelogging -Type Dword -Value 1

            Write-Host "Enable agent data upload"
            Set-GPRegistryValue -Name "Office Telemetry" -Key "HKCU\Software\Policies\Microsoft\office\16.0\osm" -ValueName EnableUpload -Type Dword -Value 1
        }
    else
    {
    Break
    }

    Write-Host 'Link the new GPO titled "Office Telemetry" to the proper OU in your environment.'



