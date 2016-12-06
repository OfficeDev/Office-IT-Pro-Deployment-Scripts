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
        <# write log#>
        $lineNum = Get-CurrentLineNumber    
        $filName = Get-CurrentFileName 
        WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Creating a new Group Policy..."

        if ($Domain) {
          New-GPO -Name $gpoName -Domain $Domain
        } else {
          New-GPO -Name $gpoName
        }
    } else {
       Write-Host "Group Policy Already Exists..."
       <# write log#>
        $lineNum = Get-CurrentLineNumber    
        $filName = Get-CurrentFileName 
        WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Group Policy Already Exists..."
    }

    #New-GPLink -Name $GpoName -Target GroupAUsers

    
    Write-Host "Configuring Group Policy '$gpoName': " -NoNewline
    <# write log#>
        $lineNum = Get-CurrentLineNumber    
        $filName = Get-CurrentFileName 
        WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Configuring Group Policy '$gpoName':"

    #keys for turning off update
    
    Set-GPRegistryValue -Name $GpoName -Key "HKLM\Software\Policies\Microsoft\office\15.0\common\officeupdate" -ValueName enableautomaticupgrade -Type DWord -Value 0 | Out-Null

    Write-Host "Done"
    <# write log#>
        $lineNum = Get-CurrentLineNumber    
        $filName = Get-CurrentFileName 
        WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Done"

    

    if (!($existingGPO)) 
    {
        Write-Host "The Group Policy will not become Active until it linked to an Active Directory Organizational Unit (OU)." `
                   "In Group Policy Management Console link the GPO titled '$gpoName' to the proper OU in your environment." -BackgroundColor Red -ForegroundColor White

        <# write log#>
        $lineNum = Get-CurrentLineNumber    
        $filName = Get-CurrentFileName 
        WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "The Group Policy will not become Active until it linked to an Active Directory Organizational Unit (OU)."
        WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "In Group Policy Management Console link the GPO titled '$gpoName' to the proper OU in your environment."
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