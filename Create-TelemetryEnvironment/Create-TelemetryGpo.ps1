Function Create-TelemetryGpo {
<#

.Synopsis
Create the Telemetry GPO on the Domain Controller

.Description
Creates a group policy that that specifies the 
Telemetry agent file share location and allows
the agent to log and upload.

.Example
Create-TelemetryGpo
The IT Pro will be prompmpted to type the 
SQL server name. The IT Pro will then be prompted
to type the version of Office being used.

#>

    Import-Module -Name grouppolicy
 
    $gpo = "Office Telemetry"
    $shareName = "TDShared"
    $SqlServer = Read-Host -Prompt 'Type the SQL Server host name'

    Write-Host "You entered $SqlServer"
        
    $officeVersion = Read-Host -Prompt 'If you are using Microsoft Office 2013 type 2013. If you are using Microsoft Office 2016 type 2016.'

    if($officeVersion -eq 2013)
    {
        Write-Host "Set the Fileshare name"
        Set-GPRegistryValue -Name "Office Telemetry" -Key "HKCU\Software\Policies\Microsoft\office\15.0\osm" -ValueName CommonFileShare -Type String -Value "\\$SqlServer\$shareName"

        Write-Host "Enable agent logging"
        Set-GPRegistryValue -Name "Office Telemetry" -Key "HKCU\Software\Policies\Microsoft\office\15.0\osm" -ValueName Enablelogging -Type Dword -Value 1

        Write-Host "Enable agent data upload"
        Set-GPRegistryValue -Name "Office Telemetry" -Key "HKCU\Software\Policies\Microsoft\office\15.0\osm" -ValueName EnableUpload -Type Dword -Value 1
    }
    else
    {
        Write-Host "Set the Fileshare name"
        Set-GPRegistryValue -Name "Office Telemetry" -Key "HKCU\Software\Policies\Microsoft\office\16.0\osm" -ValueName CommonFileShare -Type String -Value "\\$SqlServer\$shareName"

        Write-Host "Enable agent logging"
        Set-GPRegistryValue -Name "Office Telemetry" -Key "HKCU\Software\Policies\Microsoft\office\16.0\osm" -ValueName Enablelogging -Type Dword -Value 1

        Write-Host "Enable agent data upload"
        Set-GPRegistryValue -Name "Office Telemetry" -Key "HKCU\Software\Policies\Microsoft\office\16.0\osm" -ValueName EnableUpload -Type Dword -Value 1
    }

    Write-Host 'Link the new GPO titled "Office Telemetry" to the proper OU in your environment.'
}
