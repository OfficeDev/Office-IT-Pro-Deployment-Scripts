Param
(
    [Parameter(Mandatory=$true)]
    [string]$GpoName,

    [Parameter()]
    [string]$Domain = $NULL,
    
    [Parameter()]
    [string]$CommonFileShare,

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

.PARAMETER CommonFileShare
The name of the Shared Drive hosting the telemetry database.

.PARAMETER OfficeVersion
The version of office used in your environment. If a version
earlier than 2013 is used do not use this parameter.

.EXAMPLE
./Create-TelemetryGpo -GpoName "Office Telemetry" -CommonFileShare "Server1" -officeVersion 2013
A GPO named "Office Telemetry" will be created. Registry keys will be
created to enable telemetry agent logging, uploading, and the commonfileshare 
path set to \\Server1\TDShared. 

.EXAMPLE
./Create-TelemetryGpo -GpoName "Office Telemetry"
A GPO named "Office Telemetry" will be created.

#>
    Write-Host
    
    Import-Module -Name grouppolicy

    if ($Domain) {
      $existingGPO = Get-GPO -Name $gpoName -Domain $Domain -ErrorAction SilentlyContinue
    } else {
      $existingGPO = Get-GPO -Name $gpoName -ErrorAction SilentlyContinue
    }
    
    if (!($existingGPO)) 
    {
        Write-Host "Creating a new Group Policy..."

        if ($Domain) {
          New-GPO -Name $gpoName -Domain $Domain
        } else {
          New-GPO -Name $gpoName
        }
    } else {
       Write-Host "Group Policy Already Exists..."
    }

    #The same share created in Deploy-TelemetryDashboard.ps1
    $shareName = "TDShared"
    
    Write-Host "Configuring Group Policy '$gpoName': " -NoNewline

    #Office 2013
    
    Set-GPRegistryValue -Name $GpoName -Key "HKCU\Software\Policies\Microsoft\office\15.0\osm" -ValueName CommonFileShare -Type String -Value "\\$CommonFileShare\$shareName" | Out-Null
    Set-GPRegistryValue -Name $GpoName -Key "HKCU\Software\Policies\Microsoft\office\15.0\osm" -ValueName Enablelogging -Type Dword -Value 1 | Out-Null
    Set-GPRegistryValue -Name $GpoName -Key "HKCU\Software\Policies\Microsoft\office\15.0\osm" -ValueName EnableUpload -Type Dword -Value 1 | Out-Null

    #Office 2016
    Set-GPRegistryValue -Name $GpoName -Key "HKCU\Software\Policies\Microsoft\office\16.0\osm" -ValueName CommonFileShare -Type String -Value "\\$CommonFileShare\$shareName" | Out-Null
    Set-GPRegistryValue -Name $GpoName -Key "HKCU\Software\Policies\Microsoft\office\16.0\osm" -ValueName Enablelogging -Type Dword -Value 1 | Out-Null
    Set-GPRegistryValue -Name $GpoName -Key "HKCU\Software\Policies\Microsoft\office\16.0\osm" -ValueName EnableUpload -Type Dword -Value 1 | Out-Null

    Write-Host "Done"

    Write-Host
    Write-Host "The Group Policy '$gpoName' has been set to configure client to submit telemetry"
    Write-Host

    if (!($existingGPO)) 
    {
        Write-Host "The Group Policy will not become Active until it linked to an Active Directory Organizational Unit (OU)." `
                   "In Group Policy Management Console link the GPO titled '$gpoName' to the proper OU in your environment." -BackgroundColor Red -ForegroundColor White
    }

   


