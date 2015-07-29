<#
.Synopsis
Deploys the Office Telemetry Dashboard and its components

.Description
Checks for SQL installation and if not installed SQL 2014 Express
will be installed and configured. A shared folder will be created, the
telemetry processor will be installed, a new SQL database will be created,
and the telemetry agent will be configured.

.Example
./Deploy-TelemetryDashboard

#>
[string] $ErrorElevated = `
    "The script failed to run. Open an elevated command prompt window " `
    + "and try running the script again."
[string] $ErrorSqlServerNotFound = `
    "The installation failed. Try running the script again."
[string] $ErrorLcidNotFound = `
    "The language files for Office 2013 can't be found. " `
    + "Please install Telemetry Processor using the instructions " `
    + "in the Getting Started worksheet of Telemetry Dashboard."
[string] $ErrorOfficeNotFound = `
    "Office 2013 can't be found. Please verify that your computer has " `
    + "Office 2013 installed."
[string] $ErrorUnsupportedOfficeVersion = `
    "The installed version of Office isn't supported."
[string] $ErrorDpconfigNotExist = `
    "The Telemetry Processor Settings wizard can't be found. " `
    + "Please install Telemetry Processor using the instructions " `
    + "in the Getting Started worksheet of Telemetry Dashboard. "
[string] $ErrorDatabaseNameNotInRegistry = `
    "The telemetry database can't be found. Please run the " `
    + "Telemetry Processor Settings wizard again."
[string] $ErrorInstallerNotEnding = `
    "Timed out. Installation takes too long to complete."
[string] $Error32BitPowerShell64BitOS = `
    "The script failed to run. Open 64 bit version of PowerShell " `
    + "window and try running the script again."
[string] $UiMessage_Copyright `
    = "Copyright (c) 2015 Microsoft Corporation. All rights reserved."
[string] $UiMessage_Disclaimer `
    = "THIS CODE AND ANY ASSOCIATED INFORMATION ARE PROVIDED `"AS IS`" WITHOUT " `
    + "WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT " `
    + "LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS " `
    + "FOR A PARTICULAR PURPOSE.`n`n" `
    + "THE ENTIRE RISK OF USE, INABILITY TO USE, OR RESULTS FROM THE USE OF " `
    + "THIS CODE REMAINS WITH THE USER."   
[string] $UiMessage_SameVersionInstalled `
    = "The same version of Telemetry Processor is already installed. " `
    + "You need to run repair installation to continue. " `
    + "Do you want to run repair installation of Telemetry " `
    + "Processor that is already installed? Y or N"
[string] $UiMessage_PreviousVersionInstalled `
    = "Another version of Telemetry Processor exists. " `
    + "You need to remove existing version of Telemetry " `
    + "Processor to continue.`n`n" `
    + "Do you want to uninstall the version of Telemetry Processor " `
    + "that is already installed? Y or N."
[string] $UiMessage_InstallationInstruction `
    = "Installing Telemetry Processor."
[string] $UiMessage_AskForBakFilePath `
    = "Enter the full path and file name of the database backup (.bak) " `
    + " file (for example: c:\Temp\TDDB.bak)"
[string] $UiMessage_AskForBakFileAgain `
    = "The database backup file can't be found." 
[string] $UiMessage_HowToUseDashboard `
    = "`nTelemetry Dashboard deployed successfully. " `
    + "To view data in the dashboard, follow these steps:`n`n" `
    + "  1  On the Getting Started worksheet, click Connect to Database`n" `
    + "  2  Enter the SQL server and database, and then click Connect`n" `
    + "  3  Select the Documents or Solutions worksheet to see the collected data.`n"
[string] $UiMessage_NotifySQLServerDownload `
    = "Downloading Microsoft SQL server 2012 Express installer package. " `
    + "Please wait..."
[string] $UiMessage_ConfigureDatabase `
    = "Configuring the database. Please wait..." 
[string] $UiMessage_UploadData `
    = "Checking to see if data has been uploaded to the database. " `
    + "It will take a few minutes. Please wait..." 
[string] $UiMessage_RestoreDatabase `
    = "Restoring the database." 
[string] $UiMessage_CreateFolder `
    = "Creating the shared folder."
[string] $UiMessage_WriteRegFile = `
    "You can collect data from other Office 2013 client computers by " `
    + "setting the registry values that enable Telemetry Agent to collect " `
    + "and upload data.`n" `
    + "Do you want to export a .reg file? Y or N"
[string] $UiMessage_AskForReentry = `
    "Please enter either Y or N."
[string] $UiMessage_StartTelemetryProcessorSetup = `
    "Installing Telemetry Processor service..."
[string] $UiMessage_CompleteTelemetryProcessorSetup = `
    "Telemetry Processor installed successfully."
[string] $UiMessage_DpconfigCreateDatabase = `
    "To create a database, complete Telemetry Processor " `
    + "settings wizard by following the steps below:`n" `
    + "  1  On the Database Settings dialog, choose SQL server name " `
    + "from the drop down list and click Connect.`n" `
    + "  2  Enter the new database name in the SQL database text box, " `
    + "and click Create and then click Next.`n" `
    + "  3  Complete Telemetry Processor settings wizard and click Finish."
[string] $UiMessage_DpconfigFull = `
    "To create a database, complete Telemetry Processor " `
    + "settings wizard by following the steps below:`n" `
    + "  1  On the Getting started dialog, " `
    + "click Next to move to Database Settings.`n" `
    + "  2  On the Database Settings dialog, " `
    + "choose SQL server name from the drop down list and click Connect.`n" `
    + "  3  Enter the new database name in the SQL database text box, " `
    + "and click Create and then click Next.`n" `
    + "  4  On the Shared Folder dialog, Click Next.`n" `
    + "  5  Complete Telemetry Processor settings wizard and click Finish.`n"
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

#
# Global variables
#

# SQL Server 2014 Express Edition download path
[string] $InstallerUrl = "http://download.microsoft.com/download/E/A/E" `
+ "/EAE6F7FC-767A-4038-A954-49B8B05D04EB/ExpressAndTools%2064BIT/SQLEXPRWT_x64_ENU.exe"
# Name of the executable to save to the local machine
[string] $InstallerFileName = "SQLEXPRWT_x64_ENU.exe"
# Name of the database restored from the backup database file
[string] $RestoredDatabaseName = "TDDB"
# Name of the source database in the backup database file
[string] $DatabaseName = "TDDB"
# Window service name of Telemetry Processor
[string] $TelemetryProcessorServiceName = "MSDPSVC"
# Instance name used for the SQL Server installation
[string] $SuggestedInstanceName = "TDSQLEXPRESS"
# Install path
[string] $installerLocalPath=$env:TEMP + "\\" + $InstallerFileName
# Configuration file path
[string] $ConfigurationPath=$env:TEMP
# Actual configuration ini file
[string] $ConfigurationFile="$ConfigurationPath\ConfigurationFile.ini"
# Office Test value
[string] $officeTest = Get-OfficeVersion
# Get the SQL Server name
[string] $SqlServerName = Get-SqlServerName

#
# Utility functions
#

# Return the bitness of Windows.
function Test-64BitOS
{
    [Management.ManagementBaseObject] $os = Get-WMIObject win32_operatingsystem
    if ($os.OSArchitecture -match "^64")
    {
        return $true
    }
    return $false
}

# Returns the version of Office
function Get-OfficeVersion
{
$objExcel = New-Object -ComObject Excel.Application
return $objExcel.Version
}

# Tell the user to run the script in a 64-bit console
# if it is run in a 32-bit console.
function Confirm-ConsoleBitness
{
    [bool] $is64BitOS = Test-64BitOS
    
    if ($is64BitOS)
    {
        if ($PSHOME -match "SysWOW64")
        {
            write-host $Error32BitPowerShell64BitOS
            exit        
        }
    }
}

#Enable .NET
function Enable-DOTNET3
{
    #Enable .NET 3.5 if not enabled
    $feature = Get-WindowsOptionalFeature -online -FeatureName NetFx3
    if($feature.State -eq "Disabled"){
        Enable-WindowsOptionalFeature -FeatureName NetFx3 -NoRestart
    }
}

# Ask the user to answer the message again if
# the answer is not 'Y' or 'N'.
function Test-EnteredKey([string] $message)
{
    [string] $answer = $String.Empty
    do
    {
        $answer = Read-Host $message
        if ($answer -eq "Y" -or $answer -eq "N")
        {
            break
        }
        Write-Host $UiMessage_AskForReentry
        
    } while ($true)

    return $answer
}

# Inform the current status to user and continue the
# script or not based on the response.
function Read-UserResponse([string] $message)
{
    [string] $answer = Test-EnteredKey $message
    if ($answer -eq "N" -or $answer -eq "n")
    {
        Exit
    }
}

#Get existing SQL information
function Get-SqlInstance
{  

    [cmdletbinding()] 
    Param (
        [parameter(ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [Alias('__Server','DNSHostName','IPAddress')]
        [string[]]$ComputerName = $env:COMPUTERNAME
    ) 
    Process {
        ForEach ($Computer in $Computername) {
            $Computer = $computer -replace '(.*?)\..+','$1'
            Write-Verbose ("Checking {0}" -f $Computer)
            Try { 
                $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $Computer) 
                $baseKeys = "SOFTWARE\\Microsoft\\Microsoft SQL Server",
                "SOFTWARE\\Wow6432Node\\Microsoft\\Microsoft SQL Server"
                If ($reg.OpenSubKey($basekeys[0])) {
                    $regPath = $basekeys[0]
                } ElseIf ($reg.OpenSubKey($basekeys[1])) {
                    $regPath = $basekeys[1]
                } Else {
                    Continue
                }
                $regKey= $reg.OpenSubKey("$regPath")
                If ($regKey.GetSubKeyNames() -contains "Instance Names") {
                    $regKey= $reg.OpenSubKey("$regpath\\Instance Names\\SQL" ) 
                    $instances = @($regkey.GetValueNames())
                } ElseIf ($regKey.GetValueNames() -contains 'InstalledInstances') {
                    $isCluster = $False
                    $instances = $regKey.GetValue('InstalledInstances')
                } Else {
                    Continue
                }
                If ($instances.count -gt 0) { 
                    ForEach ($instance in $instances) {
                        $nodes = New-Object System.Collections.Arraylist
                        $clusterName = $Null
                        $isCluster = $False
                        $instanceValue = $regKey.GetValue($instance)
                        $instanceReg = $reg.OpenSubKey("$regpath\\$instanceValue")
                        If ($instanceReg.GetSubKeyNames() -contains "Cluster") {
                            $isCluster = $True
                            $instanceRegCluster = $instanceReg.OpenSubKey('Cluster')
                            $clusterName = $instanceRegCluster.GetValue('ClusterName')
                            $clusterReg = $reg.OpenSubKey("Cluster\\Nodes")                            
                            $clusterReg.GetSubKeyNames() | ForEach {
                                $null = $nodes.Add($clusterReg.OpenSubKey($_).GetValue('NodeName'))
                            }
                        }
                        $instanceRegSetup = $instanceReg.OpenSubKey("Setup")
                        Try {
                            $edition = $instanceRegSetup.GetValue('Edition')
                        } Catch {
                            $edition = $Null
                        }
                        Try {
                            $ErrorActionPreference = 'Stop'
                            #Get from filename to determine version
                            $servicesReg = $reg.OpenSubKey("SYSTEM\\CurrentControlSet\\Services")
                            $serviceKey = $servicesReg.GetSubKeyNames() | Where {
                                $_ -match "$instance"
                            } | Select -First 1
                            $service = $servicesReg.OpenSubKey($serviceKey).GetValue('ImagePath')
                            $file = $service -replace '^.*(\w:\\.*\\sqlservr.exe).*','$1'
                            $version = (Get-Item ("\\$Computer\$($file -replace ":","$")")).VersionInfo.ProductVersion
                        } Catch {
                            #Use potentially less accurate version from registry
                            $Version = $instanceRegSetup.GetValue('Version')
                        } Finally {
                            $ErrorActionPreference = 'Continue'
                        }
                        New-Object PSObject -Property @{
                            Computername = $Computer
                            SQLInstance = $instance
                            Edition = $edition
                            Version = $version
                            Caption = {Switch -Regex ($version) {
                                "^14" {'SQL Server 2014';Break}
                                "^11" {'SQL Server 2012';Break}
                                "^10\.5" {'SQL Server 2008 R2';Break}
                                "^10" {'SQL Server 2008';Break}
                                "^9"  {'SQL Server 2005';Break}
                                "^8"  {'SQL Server 2000';Break}
                                Default {'Unknown'}
                            }}.InvokeReturnAsIs()
                            isCluster = $isCluster
                            isClusterNode = ($nodes -contains $Computer)
                            ClusterName = $clusterName
                            ClusterNodes = ($nodes -ne $Computer)
                            FullName = {
                                If ($Instance -eq 'MSSQLSERVER') {
                                    $Computer
                                } Else {
                                    "$($Computer)\$($instance)"
                                }
                            }.InvokeReturnAsIs()
                        }
                    }
                }
            } Catch { 
                Write-Warning ("{0}: {1}" -f $Computer,$_.Exception.Message)
            }  
        }
      }
    
}

#Get the SQL Server name
function Get-SqlServerName
{
Get-SqlInstance | foreach {$_.FullName}
}

# Build the shared folder path.
# Since domain users can change the shared folder in dpconfig.exe, 
# which will then write the changed folder name in registry.
# The script gets the user selected folder name from the registry.
function Build-FileSharePath([string] $folderName)
{
    [string] $hostname = hostname
    [string] $fileSharePath = "\\" + $hostname + "\" + $folderName
    $officeTest = Get-OfficeVersion
    
    if ($officeTest -eq "16.0")
    {
        [string] $dataProcessorKey = "HKLM:\SOFTWARE\Microsoft\Office\16.0\OSM\DataProcessor"
        [string] $value = "FileShareLocation"
        $fileSharePath = Read-RegistryValue $dataProcessorKey $value
       
    return $fileSharePath
    }
    else
    {
        [string] $dataProcessorKey = "HKLM:\SOFTWARE\Microsoft\Office\15.0\OSM\DataProcessor"
        [string] $value = "FileShareLocation"
        $fileSharePath = Read-RegistryValue $dataProcessorKey $value
       
    return $fileSharePath
    }
        
}

# Return the database information stored in the data processor registry.
function Read-DataProcessorRegistry
{
    if ($officeTest -eq "16.0")
    {
    [string] $dataProcessorRegistryKey = "HKLM:\SOFTWARE\Microsoft\Office\16.0\OSM\DataProcessor"
    [string] $databaseServer = Read-RegistryValue $dataProcessorRegistryKey "DatabaseServer"
    [string] $databaseName = Read-RegistryValue $dataProcessorRegistryKey "DatabaseName"
        
    return @{ 
        DatabaseServer = $databaseServer;
        DatabaseName = $databaseName;
            }
    }
    else
    {
    [string] $dataProcessorRegistryKey = "HKLM:\SOFTWARE\Microsoft\Office\15.0\OSM\DataProcessor"
    [string] $databaseServer = Read-RegistryValue $dataProcessorRegistryKey "DatabaseServer"
    [string] $databaseName = Read-RegistryValue $dataProcessorRegistryKey "DatabaseName"
        
    return @{ 
        DatabaseServer = $databaseServer;
        DatabaseName = $databaseName;
            }
    }
}

# Print the copyright message and disclaimer.
function Print-CopyrightDisclaimer
{
    Write-Host $UiMessage_Copyright
    Write-Host $UiMessage_Disclaimer
}

# Throw an error if the script is not running in an elevated prompt.
function Check-Elevated
{
    $retval = whoami /priv | Select-String SeDebugPrivilege
    if ($retval -eq $null)
    {
        throw $ErrorElevated
    }
}

#Detect if SQL Server 2014 Express Edition is present.
function Check-SqlInstall
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
            
        }
        elseif ($wmiObject.name -match "SQL Server 2012.+Database Engine Services")
        {
            $sqlServer2012Installed = $true
        }
        elseif ($wmiObject.name -match "SQL Server 2008.+Database Engine Services")
        {
            $sqlServer2008Installed = $true
        }
        elseif ($wmiObject.name -match "SQL Server 2005.+Database Engine Services")
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
ROLE="All Features With Defaults" `
UpdateEnabled="True" `
ENU="TRUE" `
QUIET="False" `
QUIETSIMPLE="True" `
IAcceptSQLServerLicenseTerms="True" `
UpdateEnabled="True" `
ERRORREPORTING="False" `
USEMICROSOFTUPDATE="True" `
FEATURES="SQLENGINE,REPLICATION" `
UpdateSource="MU" `
HELP="False" `
INDICATEPROGRESS="False" `
X86="False" `
INSTANCENAME="TDSQLEXPRESS" `
SQMREPORTING="False" `
INSTANCEID="TDSQLEXPRESS" `
AGTSVCACCOUNT="NT AUTHORITY\NETWORK SERVICE" `
AGTSVCSTARTUPTYPE="Automatic" `
COMMFABRICPORT="0" `
COMMFABRICNETWORKLEVEL="0" `
COMMFABRICENCRYPTION="0" `
MATRIXCMBRICKCOMMPORT="0" `
SQLSVCSTARTUPTYPE="Automatic" `
FILESTREAMLEVEL="0" `
ENABLERANU="True" `
SQLCOLLATION="SQL_Latin1_General_CP1_CI_AS" `
SQLSVCACCOUNT="NT Service\MSSQL`$TDSQLEXPRESS" `
TCPENABLED="1" `
NPENABLED="1" `
BROWSERSVCSTARTUPTYPE="Automatic"
"@

New-Item $ConfigurationFile -type file -force -value $CreateIni

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

#Download SQL 2014 Express server, create the configuration file
#and install the SQL server
function Install-SqlwithIni
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



#Enable TCP/IP and set the port
function Set-TcpPort
{
    
    $TCPKey = "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\MSSQL12.TDSQLEXPRESS\MSSQLServer\SuperSocketNetLib\Tcp"
    $RegKeyIP2 = "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\MSSQL12.TDSQLEXPRESS\MSSQLServer\SuperSocketNetLib\Tcp\IP2"
    $RegKeyIPAll = "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\MSSQL12.TDSQLEXPRESS\MSSQLServer\SuperSocketNetLib\Tcp\IPAll"

    Set-ItemProperty -Path $TCPKey -Name Enabled -Value 1    
    Set-ItemProperty -Path $RegKeyIP2 -Name Enabled -Value 1
    Set-ItemProperty -Path $RegKeyIP2 -Name TcpPort -Value 1433
    Set-ItemProperty -Path $RegKeyIPAll -Name TcpPort -Value 1433

    Restart-Service "SQL Server (TDSQLEXPRESS)"
    
}

# Create a shared folder and set the permissions
function New-SharedFolder
{
    write-host $UiMessage_CreateFolder

    $ShareName = "TDShared"
    $SharedFolderPath = "$env:SystemDrive"
    
    
   if (!(Test-Path $SharedFolderPath\$ShareName))
        {
        New-Item "$env:SystemDrive\$ShareName" -Type Directory
        

        New-SMBShare –Name “TDShared” –Path “$SharedFolderPath\$ShareName”
      
        $acl = Get-Acl "$SharedFolderPath\$ShareName"
        $permission = "NT AUTHORITY\NETWORK SERVICE","FullControl","ContainerInherit,ObjectInherit","None","Allow"
        $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule $permission
        $acl.SetAccessRule($accessRule)
        $acl | Set-Acl "$SharedFolderPath\$ShareName"
        }
}


# Install the Telemetry Processor
function Install-TelemetryProcessor
{
    [string[]] $msiPath = ("C:\Program Files\Common Files\microsoft shared\OFFICE16\osmdp64.msi" `
    ,"C:\Program Files\Microsoft Office\root\VFS\ProgramFilesCommonX64\Microsoft Shared\OFFICE16\1033\osmdp64.msi" `
    ,"C:\Program Files\Common Files\microsoft shared\OFFICE16\osmdp32.msi" `
    ,"C:\Program Files\Microsoft Office\root\VFS\ProgramFilesCommonX64\Microsoft Shared\OFFICE16\1033\osmdp32.msi")
    
    if (Test-64BitOS)
    {
        if (Test-Path $msiPath[0])
        {
        Start-Process $msiPath[0] /qn -wait
        }
        elseif (Test-Path $msiPath[1])
        {
        Start-Process $msiPath[1] /qn -wait  
        } 
     }
     else
     {
        if (Test-Path $msiPath[2])
        {
        Start-Process $msiPath[2] /qn -wait
        }
        elseif (Test-Path $msiPath[3])
        {
        Start-Process $msiPath[3] /qn -wait  
        } 
     }
 } 

#Create the DataProcessor reg values
function Create-ProcessorRegData
{
    $ShareName = "TDShared"
    $databaseServer = "TDSQLEXPRESS"
    [string[]] $OSMPath = ("HKLM:\SOFTWARE\Microsoft\Office\15.0" `
    ,"HKLM:\SOFTWARE\Microsoft\Office\16.0")
    [string[]] $DataProcessorPath = ("HKLM:\SOFTWARE\Microsoft\Office\15.0\OSM" `
    ,"HKLM:\SOFTWARE\Microsoft\Office\16.0\OSM")
    $officeTest = Get-OfficeVersion
    

    if ($officeTest -eq "15.0")
    {
        New-Item -Path $OSMPath[0] -Name OSM
        New-Item -Path $DataProcessorPath[0] -Name DataProcessor
        New-ItemProperty -Path "$($DataProcessorPath[0])\DataProcessor" -Name DatabaseName -Value $DatabaseName
        New-ItemProperty -Path "$($DataProcessorPath[0])\DataProcessor" -Name DatabaseServer -Value "$env:ComputerName\$databaseServer"
        New-ItemProperty -Path "$($DataProcessorPath[0])\DataProcessor" -Name FileShareLocation -Value "$env:ComputerName\$ShareName"
    }
    else
    {
        New-Item -Path $OSMPath[1] -Name OSM
        New-Item -Path $DataProcessorPath[1] -Name DataProcessor
        New-ItemProperty -Path "$($DataProcessorPath[1])\DataProcessor" -Name DatabaseName -Value $DatabaseName
        New-ItemProperty -Path "$($DataProcessorPath[1])\DataProcessor" -Name DatabaseServer -Value "$env:ComputerName\$databaseServer"
        New-ItemProperty -Path "$($DataProcessorPath[1])\DataProcessor" -Name FileShareLocation -Value "$env:ComputerName\$ShareName"
    }
}

# Change a Windows service startup type to "Automatic".
function Set-WindowServiceToAutomatic([string] $serviceName)
{
    [Object] $service = Get-Service $serviceName
    Set-Service -InputObject $service -StartupType automatic
}


# Configure the Windows service of Telemetry Processor.
function Configure-TelemetryProcessorService([string] $database, [string] $folderName)
{
    Start-Service $TelemetryProcessorServiceName
    Set-WindowServiceToAutomatic $TelemetryProcessorServiceName
}
 
# Return the sql data root directory gotten from the registry.
function Get-SqlDataRootDirectory([string] $sqlInstance)
{
    [bool] $is64BitOS = Test-64BitOS
    [string] $key = $String.Empty
    [string] $value = "SQLDataRoot"
    [string] $sqlInstance = Get-SqlInstanceName
    [string] $dataRootDirectory = $String.Empty    
    if ($is64BitOS)
    {
        try
        {
            $key = "HKLM:\Software\Wow6432Node\Microsoft\Microsoft SQL Server\MSSQL12.$sqlInstance\Setup"
            $dataRootDirectory = Read-RegistryValue $key $value
        }
        catch
        {
            $key = "HKLM:\Software\Microsoft\Microsoft SQL Server\MSSQL12.$sqlInstance\Setup"
            $dataRootDirectory = Read-RegistryValue $key $value
        }
    }
    else
    {
        $key = "HKLM:\Software\Microsoft\Microsoft SQL Server\MSSQL12.$sqlInstance\Setup"
        $dataRootDirectory = Read-RegistryValue $key $value
    }
    $dataRootDirectory = $dataRootDirectory + "\DATA\"

    return $dataRootDirectory
}

# Read a registry value. In a 64-bit OS, read it regardless of 
# the bitness of the PowerShell console.
# Throw an exception if the value fails to be read.
function Read-RegistryValue([string] $key, [string] $value)
{
    [string] $registryValue = (Get-ItemProperty $key $value -ErrorAction Stop).$value    

    return $registryValue
}

# Return the target instance name of the SQL server to be
# used in this script.
function Get-SqlInstanceName
{
    try
    {
        Get-SQLInstance | foreach { $_.SQLInstance }
    }
    catch
    {
        return $SuggestedInstanceName
    }
}


# Return the path of dpconfig.exe.
function Get-DpconfigPath
{
    [string] $commonFilePath = $env:CommonProgramFiles
    [string] $filePath = $commonFilePath + "\" `
        + "microsoft shared\OFFICE16\dpconfig.exe"

    [bool] $fileExists = Test-Path -path $filePath
    if ($fileExists)
    {
        return $filePath;
    }
    throw $ErrorDpconfigNotExist
}

#Creates the database in the server instance
function Create-DataBase
{
    [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO')
    
    $srv = new-Object Microsoft.SqlServer.Management.Smo.Server("(local)")
    $tddb = $srv.Databases | where {$_.Name -eq 'TDDB'} 
    if (!($tddb)) 
    {
        $db = New-Object Microsoft.SqlServer.Management.Smo.Database($srv, "TDDB")
        $db.Create()  
    } 

Configure-Database

Write-Host $db.CreateDate
}


#Applies the Office Telemetry settings to the database
function Configure-Database
{
    #Import-Module SQLPS -DisableNameChecking
    Copy-Item -Path "C:\Program Files (x86)\Microsoft SQL Server\120\Tools\PowerShell\Modules\SQLPS" -Destination "C:\Windows\System32\WindowsPowerShell\v1.0\Modules"
    Invoke-Sqlcmd -ServerInstance $SqlServerName -InputFile "C:\PowerShellScripts\OfficeTelemetryDatabase.sql" -Database $DatabaseName
}


# Execute an SQL query.
function Run-SqlQuery([string] $database, [string] $query)
{
    [string] $sqlServerName = Get-SqlServerName
    $env:path = [System.Environment]::GetEnvironmentVariable("PATH", "Machine")
    [string] $command = "osql.exe"
    [string] $argument = "-E -S $sqlServerName -d $database -Q " `
        + '"' + $query + '"'

    Start-Process -FilePath $command -ArgumentList $argument -Wait
}

# Set permissions that allow SQL Serveer to get document data
# for the Office client.
function Configure-DatabasePermissions([string] $database)
{
    write-host $UiMessage_ConfigureDatabase
    
    [string] $hostname = hostname
    [string] $query = "CREATE LOGIN [NT AUTHORITY\NETWORK SERVICE] FROM WINDOWS"

    Run-SqlQuery $database $query

    $query = "CREATE USER [NT AUTHORITY\NETWORK SERVICE] FOR LOGIN " `
        + "[NT AUTHORITY\NETWORK SERVICE] WITH DEFAULT_SCHEMA=[dbo]"
    Run-SqlQuery $database $query

    $query = "EXEC sp_addrolemember 'td_telemetryprocessor', " `
        + "'NT AUTHORITY\NETWORK SERVICE'"
    Run-SqlQuery $database $query

}

# Run the task to let Telemetry Agent write data to the shared folder.
function Run-TelemetryAgentTask
{
    Start-ScheduledTask OfficeTelemetryAgentLogOn2016 \Microsoft\Office         
}

# Display instructions to the user about how to use Telemetry Dashboard.
function Show-TelemetryDashboard
{
   [string[]] $dashboardPath = ("C:\Program Files\Microsoft Office\Root\Office15\msotd.exe" `
    ,"C:\Program Files\Microsoft Office\Root\Office16\msotd.exe" `
    ,"C:\Program Files\Microsoft Office 15\root\office15\msotd.exe" `
    ,"C:\Program Files\Microsoft Office 16\root\office16\msotd.exe")

    Write-Host $UiMessage_HowToUseDashboard
    if (Test-Path $dashboardPath[0])
    {      
        Start-Process -FilePath $dashboardPath[0] /qn -Wait
    }

    elseif (Test-Path $dashboardPath[1])
    {
        Start-Process -FilePath $dashboardPath[1] /qn -wait
        
    }
    elseif (Test-Path $dashboardPath[2])
    {
        Start-Process -FilePath $dashboardPath[2] /qn -wait
        
    }
    elseif (Test-Path $dashboardPath[3])
    {
        Start-Process -FilePath $dashboardPath[3] /qn -wait
        
    }
}

# Add a registry key given its name, value and type.
function Add-RegistryKey( `
    [string] $key, `
    [string] $name, `
    [string] $value, `
    [string] $type)
{
    if (-not (Test-Path $key))
    {
        New-Item $key -Force | Out-Null
    }
    New-ItemProperty $key -Name $name -Value $value -PropertyType $type -Force | Out-Null
}

# Set the registry values to enable Telemetry Agent to upload data.
function Configure-TelemetryAgent([string] $database, [string] $folderName)
{
    [string] $key = "HKCU:\Software\Policies\Microsoft\Office\16.0\osm"
    [string] $commonFileShare = Build-FileSharePath $folderName
    Add-RegistryKey $key "CommonFileShare" $commonFileShare  "String"

    Add-RegistryKey $key "Tag1" "TAG1" "String"
    Add-RegistryKey $key "Tag2" "TAG2" "String"
    Add-RegistryKey $key "Tag3" "TAG3" "String"
    Add-RegistryKey $key "Tag4" "TAG4" "String"

    Add-RegistryKey $key "AgentInitWait" "1" "DWord"
    Add-RegistryKey $key "Enablelogging" "1" "DWord"
    Add-RegistryKey $key "EnableUpload" "1" "DWord"
    Add-RegistryKey $key "EnableFileObfuscation" "0" "DWord"
    Add-RegistryKey $key "AgentRandomDelay" "0" "DWord"
    
    Run-TelemetryAgentTask $database
}

# Give the user an option to write the .reg file.
function Write-RegFile([string] $folderName)
{
    [string] $outDirectory = $ScriptDirectory
    [string] $outPath = $outDirectory + "\agent.reg"

    [System.IO.StreamWriter] $stream = [System.IO.StreamWriter] $outPath
    [string] $commonFileShare = Build-FileSharePath $folderName
    $commonFileShare = $commonFileShare.replace("\", "\\")
    
    if ($officeTest -eq "16.0")
    {
    $stream.WriteLine("Windows Registry Editor Version 5.00")
    $stream.WriteLine("")
    $stream.WriteLine("[HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\16.0\osm]")
    $stream.WriteLine("`"CommonFileShare`"=`"$commonFileShare`"")
    $stream.WriteLine("`"Tag1`"=`"$commonFileShare`"")
    $stream.WriteLine("`"Tag2`"=`"$commonFileShare`"")
    $stream.WriteLine("`"Tag3`"=`"$commonFileShare`"")
    $stream.WriteLine("`"Tag4`"=`"$commonFileShare`"")        
    $stream.WriteLine("`"AgentInitWait`"=dword:00000001")
    $stream.WriteLine("`"Enablelogging`"=dword:00000001")
    $stream.WriteLine("`"EnableUpload`"=dword:00000001")
    $stream.WriteLine("`"EnableFileObfuscation`"=dword:00000000")
    $stream.WriteLine("`"AgentRandomDelay`"=dword:00000000")   
    $stream.Close()
    }
    else
    {
    $stream.WriteLine("Windows Registry Editor Version 5.00")
    $stream.WriteLine("")
    $stream.WriteLine("[HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\15.0\osm]")
    $stream.WriteLine("`"CommonFileShare`"=`"$commonFileShare`"")
    $stream.WriteLine("`"Tag1`"=`"$commonFileShare`"")
    $stream.WriteLine("`"Tag2`"=`"$commonFileShare`"")
    $stream.WriteLine("`"Tag3`"=`"$commonFileShare`"")
    $stream.WriteLine("`"Tag4`"=`"$commonFileShare`"")        
    $stream.WriteLine("`"AgentInitWait`"=dword:00000001")
    $stream.WriteLine("`"Enablelogging`"=dword:00000001")
    $stream.WriteLine("`"EnableUpload`"=dword:00000001")
    $stream.WriteLine("`"EnableFileObfuscation`"=dword:00000000")
    $stream.WriteLine("`"AgentRandomDelay`"=dword:00000000")   
    $stream.Close()
    }
    
    Write-Host "The .reg file has been written to $outPath"
    regedit /s $outPath
}

# Configure database, Telemetry Processor service and Telemetry Agent
# using the target database.
function Configure-DashboardComponents
{
    Create-Database
    Configure-Database $DatabaseName
    Configure-TelemetryProcessorService
    Configure-TelemetryAgent
}




# Main script flow
function Deploy-TelemetryDashboard
{

    Confirm-ConsoleBitness

    Print-CopyrightDisclaimer

    Check-Elevated

    Enable-DOTNET3

    Check-SQLInstall

    if($SqlServerName -eq $Null)
    {
        Install-SQLwithIni

        Set-TcpPort

        New-SharedFolder

        Install-TelemetryProcessor

        Create-ProcessorRegData

        Configure-DashboardComponents

        Show-TelemetryDashboard

        Write-RegFile $folderName
    }
    else
    {
        New-SharedFolder

        Install-TelemetryProcessor

        Create-ProcessorRegData

        Configure-DashboardComponents

        Show-TelemetryDashboard

        Write-RegFile $folderName
    }
}


# SIG # Begin signature block
# MIIagAYJKoZIhvcNAQcCoIIacTCCGm0CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUxK0X+531z2ovcU+PBkOALQnS
# BkWgghU/MIIEqTCCA5GgAwIBAgITMwAAAIhZDjxRH+JqZwABAAAAiDANBgkqhkiG
# 9w0BAQUFADB5MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSMw
# IQYDVQQDExpNaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBDQTAeFw0xMjA3MjYyMDUw
# NDFaFw0xMzEwMjYyMDUwNDFaMIGDMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMQ0wCwYDVQQLEwRNT1BSMR4wHAYDVQQDExVNaWNyb3NvZnQgQ29y
# cG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCzdHTQgjyH
# p5rUjrIEQoCXJS7kQc6TYzZfE/K0eJiAxih+zIoT7z03jDsJoNgUxVxe2KkdfwHB
# s5gbUHfs/up8Rc9/4SEOxYTKnw9rswk4t3TEVx6+8EioeVrfDpscmqi8yFK1DGmP
# hM5xVXv/CSC/QHc3ITB0W5Xfd8ug5cFyEgY98shVbK/B+2oWJ8j1s2Hj2c4bDx70
# 5M1MNGw+RxHnAitfFHoEB/XXPYvbZ31XPjXrbY0BQI0ah5biD3dMibo4nPuOApHb
# Ig/l0DapuDdF0Cr8lo3BYHEzpYix9sIEMIdbw9cvsnkR2ItlYqKKEWZdfn8FenOK
# H3qF5c0oENE9AgMBAAGjggEdMIIBGTATBgNVHSUEDDAKBggrBgEFBQcDAzAdBgNV
# HQ4EFgQUJls+W12WX+L3d4h/XkVTWKguW7gwDgYDVR0PAQH/BAQDAgeAMB8GA1Ud
# IwQYMBaAFMsR6MrStBZYAck3LjMWFrlMmgofMFYGA1UdHwRPME0wS6BJoEeGRWh0
# dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3RzL01pY0NvZFNp
# Z1BDQV8wOC0zMS0yMDEwLmNybDBaBggrBgEFBQcBAQROMEwwSgYIKwYBBQUHMAKG
# Pmh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljQ29kU2lnUENB
# XzA4LTMxLTIwMTAuY3J0MA0GCSqGSIb3DQEBBQUAA4IBAQAP3kBJiJHRMTejRDhp
# smor1JH7aIWuWLseDI9W+pnXypcnTOiFjnlpLOS9lj/lcGaXlTBlKa3Gyqz1D3mo
# Z79p9A+X4woPv+6WdimyItAzxv+LSa2usv2/JervJ1DA6xn4GmRqoOEXWa/xz+yB
# qInosdIUBuNqbXRSZNqWlCpcaWsf7QWZGtzoZaqIGxWVGtOkUZb9VZX4Y42fFAyx
# nn9KBP/DZq0Kr66k3mP68OrDs7Lrh9vFOK22c9J4ZOrsIVtrO9ZEIvSBUqUrQymL
# DKEqcYJCy6sbftSlp6333vdGms5DOegqU+3PQOR3iEK/RxbgpTZq76cajTo9MwT2
# JSAjMIIEwzCCA6ugAwIBAgITMwAAACs5MkjBsslI8wAAAAAAKzANBgkqhkiG9w0B
# AQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UE
# BxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEwHwYD
# VQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTIwOTA0MjExMjM0WhcN
# MTMxMjA0MjExMjM0WjCBszELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0
# b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3Jh
# dGlvbjENMAsGA1UECxMETU9QUjEnMCUGA1UECxMebkNpcGhlciBEU0UgRVNOOkMw
# RjQtMzA4Ni1ERUY4MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2
# aWNlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAprYwDgNlrlBahmuF
# n0ihHsnA7l5JB4XgcJZ8vrlfYl8GJtOLObsYIqUukq3YS4g6Gq+bg67IXjmMwjJ7
# FnjtNzg68WL7aIICaOzru0CKsf6hLDZiYHA5YGIO+8YYOG+wktZADYCmDXiLNmuG
# iiYXGP+w6026uykT5lxIjnBGNib+NDWrNOH32thc6pl9MbdNH1frfNaVDWYMHg4y
# Fz4s1YChzuv3mJEC3MFf/TiA+Dl/XWTKN1w7UVtdhV/OHhz7NL5f5ShVcFScuOx8
# AFVGWyiYKFZM4fG6CRmWgUgqMMj3MyBs52nDs9TDTs8wHjfUmFLUqSNFsq5cQUlP
# tGJokwIDAQABo4IBCTCCAQUwHQYDVR0OBBYEFKUYM1M/lWChQxbvjsav0iu6nljQ
# MB8GA1UdIwQYMBaAFCM0+NlSRnAK7UD7dvuzK7DDNbMPMFQGA1UdHwRNMEswSaBH
# oEWGQ2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3RzL01p
# Y3Jvc29mdFRpbWVTdGFtcFBDQS5jcmwwWAYIKwYBBQUHAQEETDBKMEgGCCsGAQUF
# BzAChjxodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY3Jvc29m
# dFRpbWVTdGFtcFBDQS5jcnQwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZIhvcN
# AQEFBQADggEBAH7MsHvlL77nVrXPc9uqUtEWOca0zfrX/h5ltedI85tGiAVmaiaG
# Xv6HWNzGY444gPQIRnwrc7EOv0Gqy8eqlKQ38GQ54cXV+c4HzqvkJfBprtRG4v5m
# MjzXl8UyIfruGiWgXgxCLBEzOoKD/e0ds77OkaSRJXG5q3Kwnq/kzwBiiXCpuEpQ
# jO4vImSlqOZNa5UsHHnsp6Mx2pBgkKRu/pMCDT8sJA3GaiaBUYNKELt1Y0SqaQjG
# A+vizwvtVjrs73KnCgz0ANMiuK8icrPnxJwLKKCAyuPh1zlmMOdGFxjn+oL6WQt6
# vKgN/hz/A4tjsk0SAiNPLbOFhDvioUfozxUwggW8MIIDpKADAgECAgphMyYaAAAA
# AAAxMA0GCSqGSIb3DQEBBQUAMF8xEzARBgoJkiaJk/IsZAEZFgNjb20xGTAXBgoJ
# kiaJk/IsZAEZFgltaWNyb3NvZnQxLTArBgNVBAMTJE1pY3Jvc29mdCBSb290IENl
# cnRpZmljYXRlIEF1dGhvcml0eTAeFw0xMDA4MzEyMjE5MzJaFw0yMDA4MzEyMjI5
# MzJaMHkxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQH
# EwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xIzAhBgNV
# BAMTGk1pY3Jvc29mdCBDb2RlIFNpZ25pbmcgUENBMIIBIjANBgkqhkiG9w0BAQEF
# AAOCAQ8AMIIBCgKCAQEAsnJZXBkwZL8dmmAgIEKZdlNsPhvWb8zL8epr/pcWEODf
# OnSDGrcvoDLs/97CQk4j1XIA2zVXConKriBJ9PBorE1LjaW9eUtxm0cH2v0l3511
# iM+qc0R/14Hb873yNqTJXEXcr6094CholxqnpXJzVvEXlOT9NZRyoNZ2Xx53RYOF
# OBbQc1sFumdSjaWyaS/aGQv+knQp4nYvVN0UMFn40o1i/cvJX0YxULknE+RAMM9y
# KRAoIsc3Tj2gMj2QzaE4BoVcTlaCKCoFMrdL109j59ItYvFFPeesCAD2RqGe0VuM
# JlPoeqpK8kbPNzw4nrR3XKUXno3LEY9WPMGsCV8D0wIDAQABo4IBXjCCAVowDwYD
# VR0TAQH/BAUwAwEB/zAdBgNVHQ4EFgQUyxHoytK0FlgByTcuMxYWuUyaCh8wCwYD
# VR0PBAQDAgGGMBIGCSsGAQQBgjcVAQQFAgMBAAEwIwYJKwYBBAGCNxUCBBYEFP3R
# MU7TJoqV4ZhgO6gxb6Y8vNgtMBkGCSsGAQQBgjcUAgQMHgoAUwB1AGIAQwBBMB8G
# A1UdIwQYMBaAFA6sgmBAVieX5SUT/CrhClOVWeSkMFAGA1UdHwRJMEcwRaBDoEGG
# P2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3RzL21pY3Jv
# c29mdHJvb3RjZXJ0LmNybDBUBggrBgEFBQcBAQRIMEYwRAYIKwYBBQUHMAKGOGh0
# dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljcm9zb2Z0Um9vdENl
# cnQuY3J0MA0GCSqGSIb3DQEBBQUAA4ICAQBZOT5/Jkav629AsTK1ausOL26oSffr
# X3XtTDst10OtC/7L6S0xoyPMfFCYgCFdrD0vTLqiqFac43C7uLT4ebVJcvc+6kF/
# yuEMF2nLpZwgLfoLUMRWzS3jStK8cOeoDaIDpVbguIpLV/KVQpzx8+/u44YfNDy4
# VprwUyOFKqSCHJPilAcd8uJO+IyhyugTpZFOyBvSj3KVKnFtmxr4HPBT1mfMIv9c
# Hc2ijL0nsnljVkSiUc356aNYVt2bAkVEL1/02q7UgjJu/KSVE+Traeepoiy+yCsQ
# DmWOmdv1ovoSJgllOJTxeh9Ku9HhVujQeJYYXMk1Fl/dkx1Jji2+rTREHO4QFRoA
# Xd01WyHOmMcJ7oUOjE9tDhNOPXwpSJxy0fNsysHscKNXkld9lI2gG0gDWvfPo2cK
# dKU27S0vF8jmcjcS9G+xPGeC+VKyjTMWZR4Oit0Q3mT0b85G1NMX6XnEBLTT+yzf
# H4qerAr7EydAreT54al/RrsHYEdlYEBOsELsTu2zdnnYCjQJbRyAMR/iDlTd5aH7
# 5UcQrWSY/1AWLny/BSF64pVBJ2nDk4+VyY3YmyGuDVyc8KKuhmiDDGotu3ZrAB2W
# rfIWe/YWgyS5iM9qqEcxL5rc43E91wB+YkfRzojJuBj6DnKNwaM9rwJAav9pm5bi
# EKgQtDdQCNbDPTCCBgcwggPvoAMCAQICCmEWaDQAAAAAABwwDQYJKoZIhvcNAQEF
# BQAwXzETMBEGCgmSJomT8ixkARkWA2NvbTEZMBcGCgmSJomT8ixkARkWCW1pY3Jv
# c29mdDEtMCsGA1UEAxMkTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9y
# aXR5MB4XDTA3MDQwMzEyNTMwOVoXDTIxMDQwMzEzMDMwOVowdzELMAkGA1UEBhMC
# VVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNV
# BAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEhMB8GA1UEAxMYTWljcm9zb2Z0IFRp
# bWUtU3RhbXAgUENBMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAn6Fs
# sd/bSJIqfGsuGeG94uPFmVEjUK3O3RhOJA/u0afRTK10MCAR6wfVVJUVSZQbQpKu
# mFwwJtoAa+h7veyJBw/3DgSY8InMH8szJIed8vRnHCz8e+eIHernTqOhwSNTyo36
# Rc8J0F6v0LBCBKL5pmyTZ9co3EZTsIbQ5ShGLieshk9VUgzkAyz7apCQMG6H81kw
# nfp+1pez6CGXfvjSE/MIt1NtUrRFkJ9IAEpHZhEnKWaol+TTBoFKovmEpxFHFAmC
# n4TtVXj+AZodUAiFABAwRu233iNGu8QtVJ+vHnhBMXfMm987g5OhYQK1HQ2x/Peb
# sgHOIktU//kFw8IgCwIDAQABo4IBqzCCAacwDwYDVR0TAQH/BAUwAwEB/zAdBgNV
# HQ4EFgQUIzT42VJGcArtQPt2+7MrsMM1sw8wCwYDVR0PBAQDAgGGMBAGCSsGAQQB
# gjcVAQQDAgEAMIGYBgNVHSMEgZAwgY2AFA6sgmBAVieX5SUT/CrhClOVWeSkoWOk
# YTBfMRMwEQYKCZImiZPyLGQBGRYDY29tMRkwFwYKCZImiZPyLGQBGRYJbWljcm9z
# b2Z0MS0wKwYDVQQDEyRNaWNyb3NvZnQgUm9vdCBDZXJ0aWZpY2F0ZSBBdXRob3Jp
# dHmCEHmtFqFKoKWtTHNY9AcTLmUwUAYDVR0fBEkwRzBFoEOgQYY/aHR0cDovL2Ny
# bC5taWNyb3NvZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvbWljcm9zb2Z0cm9vdGNl
# cnQuY3JsMFQGCCsGAQUFBwEBBEgwRjBEBggrBgEFBQcwAoY4aHR0cDovL3d3dy5t
# aWNyb3NvZnQuY29tL3BraS9jZXJ0cy9NaWNyb3NvZnRSb290Q2VydC5jcnQwEwYD
# VR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZIhvcNAQEFBQADggIBABCXisNcA0Q23em0
# rXfbznlRTQGxLnRxW20ME6vOvnuPuC7UEqKMbWK4VwLLTiATUJndekDiV7uvWJoc
# 4R0Bhqy7ePKL0Ow7Ae7ivo8KBciNSOLwUxXdT6uS5OeNatWAweaU8gYvhQPpkSok
# InD79vzkeJkuDfcH4nC8GE6djmsKcpW4oTmcZy3FUQ7qYlw/FpiLID/iBxoy+cwx
# SnYxPStyC8jqcD3/hQoT38IKYY7w17gX606Lf8U1K16jv+u8fQtCe9RTciHuMMq7
# eGVcWwEXChQO0toUmPU8uWZYsy0v5/mFhsxRVuidcJRsrDlM1PZ5v6oYemIp76Kb
# KTQGdxpiyT0ebR+C8AvHLLvPQ7Pl+ex9teOkqHQ1uE7FcSMSJnYLPFKMcVpGQxS8
# s7OwTWfIn0L/gHkhgJ4VMGboQhJeGsieIiHQQ+kr6bv0SMws1NgygEwmKkgkX1rq
# Vu+m3pmdyjpvvYEndAYR7nYhv5uCwSdUtrFqPYmhdmG0bqETpr+qR/ASb/2KMmyy
# /t9RyIwjyWa9nR2HEmQCPS2vWY+45CHltbDKY7R4VAXUQS5QrJSwpXirs6CWdRrZ
# kocTdSIvMqgIbqBbjCW/oO+EyiHW6x5PyZruSeD3AWVviQt9yGnI5m7qp5fOMSn/
# DsVbXNhNG6HY+i+ePy5VFmvJE6P9MYIEqzCCBKcCAQEwgZAweTELMAkGA1UEBhMC
# VVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNV
# BAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEjMCEGA1UEAxMaTWljcm9zb2Z0IENv
# ZGUgU2lnbmluZyBQQ0ECEzMAAACIWQ48UR/iamcAAQAAAIgwCQYFKw4DAhoFAKCB
# xDAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYK
# KwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQU9zxH6+upL96zzKWM/EZD7IIWNTww
# ZAYKKwYBBAGCNwIBDDFWMFSgGoAYAFQARABEAGUAcABsAG8AeQBtAGUAbgB0oTaA
# NGh0dHA6Ly9vMTUub2ZmaWNlcmVkaXIubWljcm9zb2Z0LmNvbS9yL3JsaWRUREhl
# bHBPMTUwDQYJKoZIhvcNAQEBBQAEggEAm+fi4AaPrGo8yklSLVWDzetWcZN9lAsQ
# EFWI03Zg1Ksgw8ZRJy636AX63AdumCKAA2QMboAFroq7/mYwJmHqBPmdwao/KuYm
# Fhlppr7epgCE1FPSkajlu5MR1tD3X60FGAVIRgoAfVgEn9Y+fqZhYlQPNrkwnXia
# 9CPeQdKOagzjAJqBH4dKeXmorWJMjryo7Hgq66eiXcoJYjh7wrReJRr33K8d5LAQ
# BdnYnRH9wOV4dajRQfzKcCt0oVj/kih7wufzF829lQ0NuYjc6CVSLVPHJ79Fyjta
# KOmXiXtNOiW5Q4wA86skYIGf3QpFc0R9/jcbxT49isLZYO0qInQJe6GCAigwggIk
# BgkqhkiG9w0BCQYxggIVMIICEQIBATCBjjB3MQswCQYDVQQGEwJVUzETMBEGA1UE
# CBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9z
# b2Z0IENvcnBvcmF0aW9uMSEwHwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQ
# Q0ECEzMAAAArOTJIwbLJSPMAAAAAACswCQYFKw4DAhoFAKBdMBgGCSqGSIb3DQEJ
# AzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTEyMDkyODA0MTUyNVowIwYJ
# KoZIhvcNAQkEMRYEFJZx87jLjkJl4plT2I5heHQ0+F4xMA0GCSqGSIb3DQEBBQUA
# BIIBAEWzYGpUjC3iVd+8FUnXqB4GH5ssOx8ItY7br54FNyhdr+AXkUnZ+KDMQanP
# R8mQ+qUJofCldwp3bNFTD8H6sjAdUxeE5lGYNjIxkTKfrH4vlBrH1L5+QfgVXVi0
# FF070fNMJF8Q3ZvyEoiWSOcws+eeFTKQYjMvFILkOJLV8oQb+UPKFN1NzEFBS0hJ
# fikjqdQTwVeQavecS7J/nlA8VeJTUDB5rdBvPEozQTxZFpqnPL5dBSl75ZgGGvzO
# xT/JxaX+TJekknXAhgE8lNK9yQ3O433CoIAqaKIGoHm/MBGzphpFUVcLgbJB8vaO
# ZTDrIdpKxYcrNuTC/WF79MbEznc=
# SIG # End signature block



