[string] $UiMessage_Copyright `
    = "Copyright (c) 2012 Microsoft Corporation. All rights reserved."
[string] $UiMessage_Disclaimer `
    = "THIS CODE AND ANY ASSOCIATED INFORMATION ARE PROVIDED `"AS IS`" WITHOUT " `
    + "WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT " `
    + "LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS " `
    + "FOR A PARTICULAR PURPOSE.`n`n" `
    + "THE ENTIRE RISK OF USE, INABILITY TO USE, OR RESULTS FROM THE USE OF " `
    + "THIS CODE REMAINS WITH THE USER."   
[string] $UiMessage_SqlServer2014Exists `
    = "SQL server 2014 already exists. Do you want to skip the installation" `
    + " of the SQL server 2014 Express Edition? Y or N"
[string] $UiMessage_SqlServerOtherExists `
    = "Another version of SQL server already exists, but it isn't supported " `
    + "by this script. Do you want to install SQL Server 2014 Express " `
    + "Edition? Y or N"
[string] $UiMessage_SqlServerDownloadRetry `
    = "SQL Server 2014 Express Edition can't be downloaded. Do you " `
    + "want to try again? Y or N"
[string] $UiMessage_NotifySQLServerDownload `
    = "Downloading Microsoft SQL server 2014 Express installer package. " `
    + "Please wait..."
[string] $UiMessage_AskForReentry = `
    "Please enter either Y or N."
[string] $UiMessage_StartSQLInstall = `
    "Starting SQL Server 2014 Setup process."
[string] $UiMessage_CompleteSQLInstall = `
    "SQL server installed successfully."
[string] $UiMessage_CreateConfigFile = `
    "Create configuration file."


# SQL Server 2014 Express Edition download path
[string] $InstallerUrl = "http://download.microsoft.com/download/E/A/E/EAE6F7FC-767A-4038-A954-49B8B05D04EB/Express%2064BIT/SQLEXPR_x64_ENU.exe"
# Name of the executable to save to the local machine
[string] $InstallerFileName = "SQLEXPR_x64_ENU.exe"
# Instance name used for the SQL Server installation
[string] $SuggestedInstanceName = "TDSQLEXPRESS"
#Install path
[string]$installerLocalPath=$env:TEMP + "\\" + $InstallerFileName
#Configuration file path
[string]$ConfigurationPath=$env:TEMP
#Actual configuration ini file
[string]$ConfigurationFile="$ConfigurationPath\ConfigurationFile.ini"

#Detect if SQL Server 2014 Express Edition is present.
function Check-SQLInstall
{
    [bool] $sqlServer2014Installed = $false
    [bool] $sqlServer2012Installed = $false
    [bool] $sqlServer2008Installed = $false
    [bool] $sqlServer2005Installed = $false

    [wmi[]] $wmiObjectArray = Get-WmiObject -class Win32_Product
    foreach ($wmiObject in $wmiObjectArray)
    {
        if ($wmiObject.name -match "SQL Server 2014.+Database Engine Services")
        {
            $sqlServer2014Installed = $true
            break
        }
        elseif ($wmiObject.name -match `
                    "SQL Server 2012.+Database Engine Services")
        {
            $sqlServer2012Installed = $true
        elseif ($wmiObject.name -match `
                    "SQL Server 2008.+Database Engine Services")
        {
            $sqlServer2008Installed = $true
        }
        elseif ($wmiObject.name -match "SQL Server 2005.+")
        {
            $sqlServer2005Installed = $true
        }
    }
    if ($sqlServer2014Installed)
    {
        Read-UserResponse $UiMessage_SqlServer2014Exists
    }
    elseif ($sqlServer2012Installed -or $sqlServer2008Installed -or $sqlServer2005Installed)
    {
        Read-UserResponse $UiMessage_SqlServerOtherExists

        Check-SQLInstall
    }
    else
    {
        Check-SQLInstall
    }
}
}

#Download Microsoft SQL Server 2014 Express Edition and install.
function Run-SqlServerInstaller
{
    write-host $UiMessage_StartSQLInstall
    
    [string] $installerLocalPath=$env:TEMP + "\\" + $InstallerFileName
    [System.Net.Webclient] $webClient = New-Object System.Net.WebClient
    try
    {
        write-host $UiMessage_NotifySQLServerDownload
        
        $webClient.DownloadFile($InstallerUrl, $installerLocalPath)
    }
    catch
    {
        Read-UserResponse $UiMessage_SqlServerDownloadRetry
        
        write-host $UiMessage_NotifySQLServerDownload
        
        $webClient.DownloadFile($InstallerUrl, $installerLocalPath)
    }
     
}

#Create a configuration file
function Create-ConfigurationFile {

$CreateIni = @"
[Options]
Action="Install" `
ENU="TRUE" `
QUIET="False" `
QUIETSIMPLE="True" `
IAcceptSQLServerLicenseTerms="True" `
UpdateEnabled="True" `
ROLE="All Features With Defaults" `
ERRORREPORTING="False" `
USEMICROSOFTUPDATE="True" `
UpdateSource="MU" `
HELP="False" `
INDICATEPROGRESS="False" `
X86="False" `
INSTALLSHAREDDIR="C:\Program Files\Microsoft SQL Server" `
INSTANCENAME="TDSQLEXPRESS" `
SQMREPORTING="False" `
INSTANCEID="TDSQLEXPRESS" `
INSTANCEDIR="C:\Program Files\Microsoft SQL Server" `
AGTSVCACCOUNT="NT AUTHORITY\NETWORK SERVICE" `
AGTSVCSTARTUPTYPE="Disabled" `
COMMFABRICPORT="0" `
COMMFABRICNETWORKLEVEL="0" `
COMMFABRICENCRYPTION="0" `
MATRIXCMBRICKCOMMPORT="0" `
SQLSVCSTARTUPTYPE="Automatic" `
FILESTREAMLEVEL="0" `
ENABLERANU="True" `
SQLCOLLATION="SQL_Latin1_General_CP1_CI_AS" `
SQLSVCACCOUNT="NT Service\MSSQL`$TDSQLEXPRESS" `
ADDCURRENTUSERASSQLADMIN="True" `
TCPENABLED="1" `
NPENABLED="0" `
BROWSERSVCSTARTUPTYPE="Disabled
"@

New-Item $ConfigurationFile -type file -force -value $CreateIni

}


#Download SQL 2014 Express server, create the configuration file
#and install the SQL server
function Install-SQLwithIni
{
    Run-SqlServerInstaller -wait
    
    Create-ConfigurationFile
    
    Start-Process -FilePath $installerLocalPath /ConfigurationFile=$ConfigurationFile -Wait

    [bool] $sqlServer2014Installed = $false
        [wmi[]] $wmiObjectArray = Get-WmiObject -class Win32_Product
        foreach ($wmiObject in $wmiObjectArray)
        {
            if ($wmiObject.name -match "SQL Server 2014.+Database Engine Services")
            {
            $sqlServer2014Installed = $true
            break
            }
        }
        
        if (-not $sqlServer2014Installed)
        {
        throw $ErrorSqlServerNotFound
        }
    
        write-host $UiMessage_CompleteSQLInstall

        Clear-Files
}

#Clean up files written to the client machine.
function Clear-Files
{
    [string] $installerPath = $env:temp + "\\" + $InstallerFileName
    if (Test-Path -Path $installerPath)
    {
        Remove-Item $installerPath
    }
}