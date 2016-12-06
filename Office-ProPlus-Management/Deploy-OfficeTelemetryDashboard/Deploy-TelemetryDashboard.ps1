
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
[string] $ErrorDpconfigNotExist = `
    "The Telemetry Processor Settings wizard can't be found. " `
    + "Please install Telemetry Processor using the instructions " `
    + "in the Getting Started worksheet of Telemetry Dashboard. "
[string] $Error32BitPowerShell64BitOS = `
    "The script failed to run. Open the 64 bit version of PowerShell " `
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
[string] $UiMessage_InstallationInstruction `
    = "Installing Telemetry Processor."
[string] $UiMessage_HowToUseDashboard `
    = "`nTelemetry Dashboard deployed successfully. " `
    + "To view data in the dashboard, follow these steps:`n`n" `
    + "  1  On the Getting Started worksheet, click Connect to Database`n" `
    + "  2  Enter the SQL server and database, and then click Connect`n" `
    + "  3  Select the Documents or Solutions worksheet to see the collected data.`n"
[string] $UiMessage_NotifySQLServerDownload `
    = "Downloading Microsoft SQL server 2014 Express installer package. " `
    + "Please wait..."
[string] $UiMessage_ConfigureDatabase `
    = "Configuring the database. Please wait..." 
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
[string] $InstallerFileName = "SQLEXPRWT_x64_ENU.EXE"
#Setup exe path
[string] $SetupExe = "$env:TEMP\SQLEXPRWT_x64_ENU\Setup.exe"
# Name of the database restored from the backup database file
[string] $RestoredDatabaseName = "TDDB"
# Name of the source database in the backup database file
[string] $DatabaseName = "TDDB"
# Window service name of Telemetry Processor
[string] $TelemetryProcessorServiceName = "MSDPSVC"
# Instance name used for the SQL Server installation
[string] $SuggestedInstanceName = "TDSQLEXPRESS"
# Configuration file path
[string] $ConfigurationPath="$env:TEMP\SQLEXPRWT_x64_ENU"
# Actual configuration ini file
[string] $ConfigurationFile="$ConfigurationPath\ConfigurationFile.ini"


#
# Utility functions
#

# Return the bitness of Windows.
function Test-64BitOS {
    [Management.ManagementBaseObject] $os = Get-WMIObject win32_operatingsystem
    if ($os.OSArchitecture -match "^64")
    {
        return $true
    }
    return $false
}

# Returns the version of Office
function Get-OfficeVersion {
  try {
    $objExcel = New-Object -ComObject Excel.Application
    return $objExcel.Version
  } catch [system.exception] {
     throw "Office must be installed on the local computer" 
  }
}

# Tell the user to run the script in a 64-bit console
# if it is run in a 32-bit console.
function Confirm-ConsoleBitness {
    [bool] $is64BitOS = Test-64BitOS
    
    if ($is64BitOS)
    {
        if ($PSHOME -match "SysWOW64")
        {
            write-host $Error32BitPowerShell64BitOS
            <# write log#>
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError $Error32BitPowerShell64BitOS
            exit        
        }

    Write-Host 'You are using a 64 bit OS.'
    <# write log#>
    $lineNum = Get-CurrentLineNumber    
    $filName = Get-CurrentFileName 
    WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "You are using a 64 bit OS."
    }
}

#Enable .NET
function Enable-DOTNET3 {
    #Enable .NET 3.5 if not enabled
    $feature = Get-WindowsOptionalFeature -online -FeatureName NetFx3
    if($feature.State -eq "Disabled"){
        Enable-WindowsOptionalFeature -FeatureName NetFx3 -NoRestart
    }

    Write-Host '.NET is enabled.'
    <# write log#>
    $lineNum = Get-CurrentLineNumber    
    $filName = Get-CurrentFileName 
    WriteToLogFile -LNumber $lineNum -FName $filName -ActionError ".NET is enbabled."
}

# Ask the user to answer the message again if
# the answer is not 'Y' or 'N'.
function Test-EnteredKey([string] $message) {
    [string] $answer = $String.Empty
    do
    {
        $answer = Read-Host $message
        if ($answer -eq "Y" -or $answer -eq "N")
        {
            break
        }
        Write-Host $UiMessage_AskForReentry
        <# write log#>
        $lineNum = Get-CurrentLineNumber    
        $filName = Get-CurrentFileName 
        WriteToLogFile -LNumber $lineNum -FName $filName -ActionError $UiMessage_AskForReentry
        
    } while ($true)

    return $answer
}

# Inform the current status to user and continue the
# script or not based on the response.
function Read-UserResponse([string] $message) {
    [string] $answer = Test-EnteredKey $message
    if ($answer -eq "N" -or $answer -eq "n")
    {
        Exit
    }
}

#Get existing SQL information
function Get-SqlInstance {  

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
function Get-SqlServerName {
Get-SqlInstance | foreach {$_.FullName}
}

function Get-SqlVersion {
#NB: replace "." with instance name (e.g. ".\sqlexpress" or "axlive\sql06")
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.SMO") | Out-Null
$sqlVersion = New-Object -typeName Microsoft.SqlServer.Management.Smo.Server(".") | select version
return $sqlVersion
}

# Build the shared folder path.
# The script gets the user selected folder name from the registry.
function Build-FileSharePath([string] $folderName) {
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
function Read-DataProcessorRegistry {
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

# Throw an error if the script is not running in an elevated prompt.
function Check-Elevated {
    $retval = whoami /priv | Select-String SeDebugPrivilege
    if ($retval -eq $null)
    {
        throw $ErrorElevated
    }
    else
    {
    Write-Host 'The script is running in an elevated prompt.'
    <# write log#>
        $lineNum = Get-CurrentLineNumber    
        $filName = Get-CurrentFileName 
        WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "The script is running in an elevated prompt."
    }
}

#Detect if SQL Server 2014 Express Edition is present.
function Check-SqlInstall {
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
function Run-SqlServerInstaller {
    write-host $UiMessage_StartSQLInstall
    <# write log#>
    $lineNum = Get-CurrentLineNumber    
    $filName = Get-CurrentFileName 
    WriteToLogFile -LNumber $lineNum -FName $filName -ActionError $UiMessage_StartSQLInstall
    
    [string] $installerLocalPath=$env:TEMP + "\\" + $InstallerFileName
    [System.Net.Webclient] $webClient = New-Object System.Net.WebClient
    try
    {
        write-host $UiMessage_NotifySQLServerDownload
        <# write log#>
        $lineNum = Get-CurrentLineNumber    
        $filName = Get-CurrentFileName 
        WriteToLogFile -LNumber $lineNum -FName $filName -ActionError $UiMessage_NotifySQLServerDownload
        
        $webClient.DownloadFile($InstallerUrl, $installerLocalPath)
    }
    catch
    {
        Read-UserResponse $UiMessage_SqlServerDownloadRetry
        
        write-host $UiMessage_NotifySQLServerDownload
        <# write log#>
        $lineNum = Get-CurrentLineNumber    
        $filName = Get-CurrentFileName 
        WriteToLogFile -LNumber $lineNum -FName $filName -ActionError $UiMessage_NotifySQLServerDownload
        
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
function Clear-Files {
    [string] $installerPath = $env:TEMP + "\\" + $InstallerFileName
    if (Test-Path -Path $installerPath)
    {
        Remove-Item $installerPath
    }
}

#Download SQL 2014 Express server, create the configuration file
#and install the SQL server
function Install-SqlwithIni {
    Run-SqlServerInstaller -wait
    
    Push-location $env:TEMP

    .\SQLEXPRWT_x64_ENU.exe /Q | Out-Null

    Pop-Location

    Create-ConfigurationFile

    Push-Location $ConfigurationPath 
    
    start-process "SETUP.EXE" /ConfigurationFile=$ConfigurationFile -Wait

    Pop-Location

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
        <# write log#>
        $lineNum = Get-CurrentLineNumber    
        $filName = Get-CurrentFileName 
        WriteToLogFile -LNumber $lineNum -FName $filName -ActionError $UiMessage_CompleteSQLInstall

        Clear-Files
}

#Enable TCP/IP and set the port
function Set-TcpPort {

    $SqlInstanceName = Get-SqlInstance | foreach { $_.SQLInstance }
    
    $TCPKey = "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\MSSQL12.$SqlInstanceName\MSSQLServer\SuperSocketNetLib\Tcp"
    $RegKeyIP2 = "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\MSSQL12.$SqlInstanceName\MSSQLServer\SuperSocketNetLib\Tcp\IP2"
    $RegKeyIPAll = "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\MSSQL12.$SqlInstanceName\MSSQLServer\SuperSocketNetLib\Tcp\IPAll"

    Set-ItemProperty -Path $TCPKey -Name Enabled -Value 1    
    Set-ItemProperty -Path $RegKeyIP2 -Name Enabled -Value 1
    Set-ItemProperty -Path $RegKeyIP2 -Name TcpPort -Value 1433
    Set-ItemProperty -Path $RegKeyIPAll -Name TcpPort -Value 1433

    Restart-Service -Name "MSSQL`$$SqlInstanceName" -WarningAction SilentlyContinue
    
}

# Create a shared folder and set the permissions
function New-SharedFolder {
    write-host $UiMessage_CreateFolder
    <# write log#>
    $lineNum = Get-CurrentLineNumber    
    $filName = Get-CurrentFileName 
    WriteToLogFile -LNumber $lineNum -FName $filName -ActionError $UiMessage_CreateFolder

    $ShareName = "TDShared"
    $SharedFolderPath = "$env:SystemDrive"
    
    if (!(Test-Path $SharedFolderPath\$ShareName))
    {
        New-Item "$env:SystemDrive\$ShareName" -Type Directory
        
        net share 'TDShared=C:\TDShared' '/Grant:Authenticated Users,Change'
      
        $acl = Get-Acl "$SharedFolderPath\$ShareName"
        $permission = "NT AUTHORITY\NETWORK SERVICE","FullControl","ContainerInherit,ObjectInherit","None","Allow"
        $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule $permission
        $acl.SetAccessRule($accessRule)
        $acl | Set-Acl "$SharedFolderPath\$ShareName"
    }

    $ShareName = "TelemetryAgent"
    $SharedFolderPath = "$env:SystemDrive"

    if (!(Test-Path $SharedFolderPath\$ShareName))
    {
        New-Item "$env:SystemDrive\$ShareName" -Type Directory
        
        net share 'TelemetryAgent=C:\TelemetryAgent'
    }

    if (!(Test-Path "C:\TelemetryAgent\osmia32.msi"))
    {
       Copy-Item -Path "$PSScriptRoot\osmia32.msi" -Destination "C:\TelemetryAgent" -Force
    }
    if (!(Test-Path "C:\TelemetryAgent\osmia64.msi"))
    {
        Copy-Item -Path "$PSScriptRoot\osmia64.msi" -Destination "C:\TelemetryAgent" -Force 
    }
}

# Install the Telemetry Processor
function Install-TelemetryProcessor {

    if(Test-64BitOS)
    {
            [string[]] $msiPath = ("C:\Program Files\Microsoft Office 15\root\vfs\ProgramFilesCommonX86\Microsoft Shared\OFFICE15\1033\osmdp64.msi",
                                   "C:\Program Files\Microsoft Office 15\root\vfs\ProgramFilesCommonX64\Microsoft Shared\OFFICE15\1033\osmdp64.msi",
                                   "C:\Program Files\Microsoft Office\root\VFS\ProgramFilesCommonX86\Microsoft Shared\OFFICE15\1033\osmdp64.msi",
                                   "C:\Program Files\Microsoft Office\root\VFS\ProgramFilesCommonX64\Microsoft Shared\OFFICE15\1033\osmdp64.msi",
                                   "C:\Program Files\Microsoft Office\root\vfs\ProgramFilesCommonX86\Microsoft Shared\OFFICE16\1033\osmdp64.msi",
                                   "C:\Program Files\Microsoft Office\root\vfs\ProgramFilesCommonX64\Microsoft Shared\OFFICE16\1033\osmdp64.msi",
                                   "C:\Program Files (x86)\Microsoft Office\root\VFS\ProgramFilesCommonX86\Microsoft Shared\OFFICE15\1033\osmdp64.msi",                               
                                   "C:\Program Files (x86)\Microsoft Office\root\VFS\ProgramFilesCommonX86\Microsoft Shared\OFFICE16\1033\osmdp64.msi",
                                   "C:\Program Files (x86)\Common Files\microsoft shared\OFFICE15\osmdp64.msi",
                                   "C:\Program Files (x86)\Common Files\microsoft shared\OFFICE16\osmdp64.msi")
     
    }
    else
    {
        [string[]] $msiPath = ("C:\Program Files\Microsoft Office 15\root\vfs\ProgramFilesCommonX86\Microsoft Shared\OFFICE15\1033\osmdp32.msi",
                               "C:\Program Files\Microsoft Office 15\root\vfs\ProgramFilesCommonX64\Microsoft Shared\OFFICE15\1033\osmdp32.msi",
                               "C:\Program Files\Microsoft Office\root\VFS\ProgramFilesCommonX86\Microsoft Shared\OFFICE15\1033\osmdp32.msi",
                               "C:\Program Files\Microsoft Office\root\VFS\ProgramFilesCommonX64\Microsoft Shared\OFFICE15\1033\osmdp32.msi",
                               "C:\Program Files\Microsoft Office\root\vfs\ProgramFilesCommonX86\Microsoft Shared\OFFICE16\1033\osmdp32.msi",
                               "C:\Program Files\Microsoft Office\root\vfs\ProgramFilesCommonX64\Microsoft Shared\OFFICE16\1033\osmdp32.msi",
                               "C:\Program Files (x86)\Microsoft Office\root\VFS\ProgramFilesCommonX86\Microsoft Shared\OFFICE15\1033\osmdp32.msi",                               
                               "C:\Program Files (x86)\Microsoft Office\root\VFS\ProgramFilesCommonX86\Microsoft Shared\OFFICE16\1033\osmdp32.msi",
                               "C:\Program Files (x86)\Common Files\microsoft shared\OFFICE15\osmdp32.msi",
                               "C:\Program Files (x86)\Common Files\microsoft shared\OFFICE16\osmdp32.msi")

    }
            
    foreach ($path in $msiPath) 
    {
        if (Test-Path $path)
        {
            Start-Process $path /qn -wait                
        }
    }
} 

#Create the DataProcessor reg values
function Create-ProcessorRegData {
    $SqlInstanceName = Get-SqlInstance | foreach { $_.SQLInstance }
    $ShareName = "TDShared"
    $databaseServer = $SqlInstanceName
    [string[]] $OSMPath = ("HKLM:\SOFTWARE\Microsoft\Office\15.0" `
    ,"HKLM:\SOFTWARE\Microsoft\Office\16.0")
    [string[]] $DataProcessorPath = ("HKLM:\SOFTWARE\Microsoft\Office\15.0\OSM",
                                     "HKLM:\SOFTWARE\Microsoft\Office\16.0\OSM")
    $officeTest = Get-OfficeVersion
    

    if ($officeTest -eq "15.0")
    {
        New-Item -Path $OSMPath[0] -Name OSM -ErrorAction SilentlyContinue
        New-Item -Path $DataProcessorPath[0] -Name DataProcessor -ErrorAction SilentlyContinue
        New-ItemProperty -Path "$($DataProcessorPath[0])\DataProcessor" -Name DatabaseName -Value $DatabaseName -ErrorAction SilentlyContinue
        New-ItemProperty -Path "$($DataProcessorPath[0])\DataProcessor" -Name DatabaseServer -Value "$env:ComputerName\$databaseServer" -ErrorAction SilentlyContinue
        New-ItemProperty -Path "$($DataProcessorPath[0])\DataProcessor" -Name FileShareLocation -Value "\\$env:ComputerName\$ShareName" -ErrorAction SilentlyContinue
    }
    else
    {
        New-Item -Path $OSMPath[1] -Name OSM -ErrorAction SilentlyContinue
        New-Item -Path $DataProcessorPath[1] -Name DataProcessor -ErrorAction SilentlyContinue
        New-ItemProperty -Path "$($DataProcessorPath[1])\DataProcessor" -Name DatabaseName -Value $DatabaseName -ErrorAction SilentlyContinue
        New-ItemProperty -Path "$($DataProcessorPath[1])\DataProcessor" -Name DatabaseServer -Value "$env:ComputerName\$databaseServer" -ErrorAction SilentlyContinue
        New-ItemProperty -Path "$($DataProcessorPath[1])\DataProcessor" -Name FileShareLocation -Value "\\$env:ComputerName\$ShareName" -ErrorAction SilentlyContinue
    }
}

# Change a Windows service startup type to "Automatic".
function Set-WindowServiceToAutomatic([string] $serviceName) {
    [Object] $service = Get-Service $serviceName
    Set-Service -InputObject $service -StartupType automatic
}


# Configure the Windows service of Telemetry Processor.
function Configure-TelemetryProcessorService([string] $database, [string] $folderName) {
    Start-Service $TelemetryProcessorServiceName
    Set-WindowServiceToAutomatic $TelemetryProcessorServiceName
}
 
# Return the sql data root directory gotten from the registry.
function Get-SqlDataRootDirectory([string] $sqlInstance) {
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
function Read-RegistryValue([string] $key, [string] $value) {
    [string] $registryValue = (Get-ItemProperty $key $value -ErrorAction Stop).$value    

    return $registryValue
}

# Return the target instance name of the SQL server to be
# used in this script.
function Get-SqlInstanceName {

    $SqlInstanceName = Get-SQLInstance | foreach { $_.SQLInstance }
    
    Write-Host $SqlInstanceName
    <# write log#>
    $lineNum = Get-CurrentLineNumber    
    $filName = Get-CurrentFileName 
    WriteToLogFile -LNumber $lineNum -FName $filName -ActionError $SqlInstanceName
}


# Return the path of dpconfig.exe.
function Get-DpconfigPath {
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

#Copy the SQLPS folder
function Copy-Sqlps {

    $SqlVersion = Get-SqlVersion

    if ($SqlVersion -match '8')
    {
    $sqlpsPath = "C:\Program Files (x86)\Microsoft SQL Server\80\Tools\PowerShell\Modules\SQLPS\*"    
    }
    elseif ($SqlVersion -match '9')
    {
    $sqlpsPath = "C:\Program Files (x86)\Microsoft SQL Server\90\Tools\PowerShell\Modules\SQLPS\*"
    }
    elseif ($SqlVersion -match '10')
    {
    $sqlpsPath = "C:\Program Files (x86)\Microsoft SQL Server\100\Tools\PowerShell\Modules\SQLPS\*"
    }
    elseif ($SqlVersion -match '11')
    {
    $sqlpsPath = "C:\Program Files (x86)\Microsoft SQL Server\110\Tools\PowerShell\Modules\SQLPS\*"
    }
    elseif ($SqlVersion -match '12')
    {
    $sqlpsPath = "C:\Program Files (x86)\Microsoft SQL Server\120\Tools\PowerShell\Modules\SQLPS\*"
    }

    $destinationPath = "$env:windir\System32\WindowsPowerShell\v1.0\Modules\SQLPS"

    if(!(Test-Path -Path $destinationPath))
    {
        Copy-Item -Path $sqlpsPath -Destination $destinationPath
    }
}

#Creates the database in the server instance
function Create-DataBase {

    Import-Module SQLPS -DisableNameChecking
        
    $srv = new-Object Microsoft.SqlServer.Management.Smo.Server("(local)")
    $tddb = $srv.Databases | where {$_.Name -eq 'TDDB'} 

    if (!($tddb)) 
    {
        $db = New-Object Microsoft.SqlServer.Management.Smo.Database($srv, "TDDB")
        $db.Create()  
    }

    Configure-Database

    Configure-DatabasePermissions 'TDDB'

    Set-Location $env:SystemDrive

    Write-Host $db.CreateDate
}


#Applies the Office Telemetry settings to the database
function Configure-Database {    
    Invoke-Sqlcmd -ServerInstance $SqlServerName -InputFile "$PSScriptRoot\OfficeTelemetryDatabase.sql" -Database $DatabaseName -ErrorAction SilentlyContinue
}


# Execute an SQL query.
function Run-SqlQuery([string] $database, [string] $query) {
    [string] $sqlServerName = Get-SqlServerName
    $env:path = [System.Environment]::GetEnvironmentVariable("PATH", "Machine")
    [string] $command = "osql.exe"
    [string] $argument = "-E -S $sqlServerName -d $database -Q " `
        + '"' + $query + '"'

    Start-Process -FilePath $command -ArgumentList $argument -Wait
}

# Set permissions that allow SQL Serveer to get document data
# for the Office client.
function Configure-DatabasePermissions([string] $database) {
    write-host $UiMessage_ConfigureDatabase
    <# write log#>
    $lineNum = Get-CurrentLineNumber    
    $filName = Get-CurrentFileName 
    WriteToLogFile -LNumber $lineNum -FName $filName -ActionError $UiMessage_ConfigureDatabase
    
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
function Run-TelemetryAgentTask {
   $logon2016 = Get-ScheduledTask | Where { $_.TaskName -eq 'OfficeTelemetryAgentLogOn2016' }
   $logon = Get-ScheduledTask | Where { $_.TaskName -eq 'OfficeTelemetryAgentLogOn' }

   if ($logon2016) {
       Start-ScheduledTask Microsoft\Office\OfficeTelemetryAgentLogOn2016 
   }
   if ($logon) {
       Start-ScheduledTask Microsoft\Office\OfficeTelemetryAgentLogOn 
   }
}

# Display instructions to the user about how to use Telemetry Dashboard.
function Show-TelemetryDashboard {
   [string[]] $dashboardPath = ("C:\Program Files\Microsoft Office\Root\Office15\msotd.exe",
                                "C:\Program Files\Microsoft Office\Root\Office16\msotd.exe",
                                "C:\Program Files\Microsoft Office 15\root\office15\msotd.exe",
                                "C:\Program Files\Microsoft Office 16\root\office16\msotd.exe",
                                "C:\Program Files (x86)\Microsoft Office\root\Office16\msotd.exe",
                                "C:\Program Files (x86)\Microsoft Office 15\root\Office15\msotd.exe")

    Write-Host $UiMessage_HowToUseDashboard
    <# write log#>
    $lineNum = Get-CurrentLineNumber    
    $filName = Get-CurrentFileName 
    WriteToLogFile -LNumber $lineNum -FName $filName -ActionError $UiMessage_HowToUseDashboard
    foreach($path in $dashboardPath)
    {
        if(Test-Path -Path $path)
        {
            Start-Process -FilePath $path /qn -Wait
        }
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
function Configure-TelemetryAgent([string] $database, [string] $folderName) {
    [string] $key = "HKCU:\Software\Policies\Microsoft\Office\16.0\osm"
    [string] $commonFileShare = Build-FileSharePath $folderName
    Add-RegistryKey $key "CommonFileShare" $commonFileShare  "String"

    Add-RegistryKey $key "Tag1" "$commonFileShare" "String"
    Add-RegistryKey $key "Tag2" "$commonFileShare" "String"
    Add-RegistryKey $key "Tag3" "$commonFileShare" "String"
    Add-RegistryKey $key "Tag4" "$commonFileShare" "String"

    Add-RegistryKey $key "AgentInitWait" "1" "DWord"
    Add-RegistryKey $key "Enablelogging" "1" "DWord"
    Add-RegistryKey $key "EnableUpload" "1" "DWord"
    Add-RegistryKey $key "EnableFileObfuscation" "0" "DWord"
    Add-RegistryKey $key "AgentRandomDelay" "0" "DWord"
    
    Run-TelemetryAgentTask $database
}

# Give the user an option to write the .reg file.
function Write-RegFile([string] $folderName) {
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
function Configure-DashboardComponents {

    Copy-Sqlps
    
    Create-Database

    Configure-TelemetryProcessorService

    Configure-TelemetryAgent
}

# Main script flow
Confirm-ConsoleBitness

Check-Elevated

Enable-DOTNET3

Check-SQLInstall

Get-SqlInstance | Out-Null

$SqlServerName = Get-SqlServerName

    if($SqlServerName -eq $null) {

        Install-SQLwithIni

        Set-TcpPort

        New-SharedFolder

        Install-TelemetryProcessor

        Create-ProcessorRegData

        Configure-DashboardComponents

        Show-TelemetryDashboard

        Write-RegFile $folderName
    }
    else {
        New-SharedFolder

        Install-TelemetryProcessor

        Create-ProcessorRegData

        Configure-DashboardComponents

        Show-TelemetryDashboard

        Write-RegFile $folderName
    }



function Get-CurrentLineNumber {
    $MyInvocation.ScriptLineNumber
}

function Get-CurrentFileName{
    $MyInvocation.ScriptName.Substring($MyInvocation.ScriptName.LastIndexOf("\")+1)
}

function Get-CurrentFunctionName {
    (Get-Variable MyInvocation -Scope 1).Value.MyCommand.Name;
}

Function WriteToLogFile() {
    param( 
        [Parameter(Mandatory=$true)]
        [string]$LNumber,
        [Parameter(Mandatory=$true)]
        [string]$FName,
        [Parameter(Mandatory=$true)]
        [string]$ActionError
    )
    try{
        $headerString = "Time".PadRight(30, ' ') + "Line Number".PadRight(15,' ') + "FileName".PadRight(60,' ') + "Action"
        $stringToWrite = $(Get-Date -Format G).PadRight(30, ' ') + $($LNumber).PadRight(15, ' ') + $($FName).PadRight(60,' ') + $ActionError

        #check if file exists, create if it doesn't
        $getCurrentDatePath = "C:\Windows\Temp\" + (Get-Date -Format u).Substring(0,10)+"OfficeAutoScriptLog.txt"
        if(Test-Path $getCurrentDatePath){#if exists, append
             Add-Content $getCurrentDatePath $stringToWrite
        }
        else{#if not exists, create new
             Add-Content $getCurrentDatePath $headerString
             Add-Content $getCurrentDatePath $stringToWrite
        }
    } catch [Exception]{
        Write-Host $_
    }
}