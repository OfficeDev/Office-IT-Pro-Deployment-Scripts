﻿[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true, Position=0)]
    [System.Management.Automation.PSCredential]$Credentials=$NULL
)

Function Restrict-ExchangeOnlineMobileAccess {

<#
.Synopsis
Allows only Outlook Mobile Device access to Exchange Online Mailboxes

.DESCRIPTION
This function will connect to Exchange Online, allow Outlook Mobile App access, and disable all other Mobile Device access to all Mailboxes in an Exchange Online Tenant

.NOTES   
Name: RestrictMobileAccess
Version: 0.1.0
DateCreated: 2015-09-16
DateUpdated: 2015-09-17

.LINK
https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts

.PARAMETER Credentials
This parameter is for the Office 365 Admin credentials that have Exchange Online administrative access
to the Office 365 Tenant.  The account must be in the 'Recipient Management' or 'Organization Managment' role.
The username must be your Office 365 username.

.EXAMPLE
.\Restrict-ExchangeOnlineMobileAccess

Description:
Running the script with no parameters will prompt you to provide Office 365 credentials

.EXAMPLE

$userName = "admin@tenant.onmicrosoft.com"
$securedPassword = ConvertTo-SecureString "PlainTextPassword" -AsPlainText -Force
$credentials = New-Object System.Management.Automation.PSCredential ($userName, $securedPassword)

Disable-ExchangeOnlineMobileAccess -Credentials $credentials

Description:
In this example you can create provide the username and password with a prompt

#>

[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true, Position=0)]
    [System.Management.Automation.PSCredential]$Credentials=$NULL
)

begin {
 
}

process {

   if (!($Credentials)) {
      $Credentials = Get-Credential
   }

   if ($Credentials) {
       $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Credentials -Authentication Basic -AllowRedirection -WarningAction SilentlyContinue -ErrorAction Stop

       Import-PSSession $Session -ErrorAction Stop -WarningAction SilentlyContinue -AllowClobber | Out-Null

       Write-Host
       Write-Host "Disabling OWA for Mobile Devices: " -NoNewline
        #write log
        $lineNum = Get-CurrentLineNumber    
        $filName = Get-CurrentFileName 
        WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Disabling OWA for Mobile Devices: "
       Get-Mailbox | Set-CasMailbox -OWAforDevicesEnabled $False -WarningAction SilentlyContinue
       Write-Host "Complete"

       Write-Host "Creating Access Rule to explicitly allow Outlook Mobile Access: " -NoNewline
        #write log
        $lineNum = Get-CurrentLineNumber    
        $filName = Get-CurrentFileName 
        WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Complete, Creating Access Rule to explicitly allow Outlook Mobile Access:"

       $ruleExists = $false
       $existingRules = Get-ActiveSyncDeviceAccessRule | where { $_.QueryString -eq "Outlook for iOS and Android" -and $_.AccessLevel -eq "Allow" -and $_.Characteristic -eq "DeviceModel" }
       if ($existingRules) {
          $ruleExists = $true
       }

       if (!($ruleExists)) {
          New-ActiveSyncDeviceAccessRule -Characteristic DeviceModel -QueryString "Outlook for iOS and Android" -AccessLevel Allow | Out-Null
          Write-Host "Complete"
            #write log
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Complete"
       } else {
          Write-Host "Already Exists"
            #write log
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Already Exists"
       }

       Write-Host "Setting ActiveSync Mobile Device Access to 'Block': " -NoNewline
        #write log
        $lineNum = Get-CurrentLineNumber    
        $filName = Get-CurrentFileName 
        WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Setting ActiveSync Mobile Device Access to 'Block': "

       $asOrgSettings = Get-ActiveSyncOrganizationSettings

       if ($asOrgSettings.DefaultAccessLevel -ne "Block") {
           Set-ActiveSyncOrganizationSettings -DefaultAccessLevel Block -WarningAction SilentlyContinue | Out-Null
           Write-Host "Complete"
            #write log
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Complete"
       } else {
          Write-Host "Already Set"
            #write log
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Already Set"
       }

       #Remove-PSSession $Session
   }
}

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


Restrict-ExchangeOnlineMobileAccess -Credentials $Credentials
