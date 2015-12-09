Param
(
    [Parameter(Mandatory=$true)]
    [string]$GpoName,

    [Parameter()]
    [string]$Domain = $NULL

)

<#

.SYNOPSIS
Create the Suppress MS Office 2016 Update GPO on the Domain Controller

.DESCRIPTION
Creates a group policy that that prevents Office 2013
from upgrading to Office 2016

.PARAMETER GpoName
Required, The name of the GPO to be created.
.PARAMETER Domain
Optional, The name of the domain the GPO will
belong to


.EXAMPLE
. C:\Users\name\Documents\PreventOfficeUpgrade.ps1 -GpoName SuppressMSOffice2016 -Domain er.mobap.com
Will create the GPO named "SuppressMSOffice2016" on the domain "er.mobap.com"

.EXAMPLE
. C:\Users\emsadmin\Documents\PreventOfficeUpgrade.ps1 -GpoName SuppressMSOffice2016
Will create the GPO named "SuppressMSOffice2016" no domain will be assigned
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

    #New-GPLink -Name $GpoName -Target GroupAUsers

    
    Write-Host "Configuring Group Policy '$gpoName': " -NoNewline

    #keys for turning off update
    
    Set-GPRegistryValue -Name $GpoName -Key "HKLM\Software\Policies\Microsoft\office\15.0\common\officeupdate" -ValueName enableautomaticupgrade -Type DWord -Value 0 | Out-Null

    Write-Host "Done"

    

    if (!($existingGPO)) 
    {
        Write-Host "The Group Policy will not become Active until it linked to an Active Directory Organizational Unit (OU)." `
                   "In Group Policy Management Console link the GPO titled '$gpoName' to the proper OU in your environment." -BackgroundColor Red -ForegroundColor White
    }
