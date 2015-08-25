[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true, Position=0)]
    [System.Management.Automation.PSCredential]$Credentials=$NULL
)

Function Disable-ExchangeOnlineMobileAccess {
<#
.Synopsis
Disables Mobile Devices access to Exchange Online Mailboxes

.DESCRIPTION
This function will connect to Exchange Online and disable Mobile Device access to all Mailboxes in an Exchange Online Tenant

.NOTES   
Name: Get-OfficeVersion
Version: 0.1.0
DateCreated: 2015-08-24
DateUpdated: 2015-08-24

.LINK
https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts

.PARAMETER Credentials
The computer or list of computers from which to query 

.EXAMPLE
Disable-ExchangeOnlineMobileAccess -Credentials Get-Credential

Description:


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
       $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Credentials -Authentication Basic -AllowRedirection -WarningAction SilentlyContinue

       Import-PSSession $Session | Out-Null

       Write-Host
       Write-Host "Disabling Mobile Mailbox access for all Mailboxes..."

       Get-Mailbox | Set-CasMailbox -ActiveSyncEnabled $False -OWAforDevicesEnabled $False

       Write-Host "Disabling Mobile Mailbox access for all Mailboxes: Complete"

       Remove-PSSession $Session
   }
}

}

Disable-ExchangeOnlineMobileAccess -Credentials $Credentials
