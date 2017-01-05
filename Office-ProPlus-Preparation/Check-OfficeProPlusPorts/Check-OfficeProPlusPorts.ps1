Function Check-OfficeProPlusPorts {
<#
.Synopsis
Checks the availability of the various remote resources needed to install Office 365

.DESCRIPTION
Checks the availability of the various remote resources needed to install Office 365

.EXAMPLE
Check-OfficeProPlusPorts

.LINK
https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts

.NOTES   
Name: Check-OfficeProPlusPorts
Version: 1.0.0
DateCreated: 2015-11-10
DateUpdated: 2015-11-12



#>
[CmdletBinding(SupportsShouldProcess=$true)]
param(

)

begin {
    $defaultDisplaySet = 'Name', 'Host', 'Port', 'Status'

    $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet(‘DefaultDisplayPropertySet’,[string[]]$defaultDisplaySet)
    $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
}


process {

    $results = new-object PSObject[] 0;

    $results += New-Object PSObject -Property @{Name = "Renew Product Key"; Host = "activation.sls.microsoft.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Validate Certificates"; Host = "crl.microsoft.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Validate Certificates"; Host = "crl.microsoft.com"; Port = 80; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Identity Configuration Services"; Host = "odc.officeapps.live.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Identity Configuration Services"; Host = "clientconfig.microsoftonline-p.net"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Office Licensing Service"; Host = "ols.officeapps.live.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Redirection Services"; Host = "office15client.microsoft.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Installation/Update Content"; Host = "officecdn.microsoft.com"; Port = 80; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Online Help Services"; Host = "go.microsoft.com"; Port = 80; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Office client only"; Host = "ocws.officeapps.live.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "In app help"; Host = "ocsa.officeapps.live.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Bing image search"; Host = "insertmedia.bing.office.net"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Roaming Services"; Host = "ea-roaming.officeapps.live.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Configuration Services"; Host = "ea-roaming.officeapps.live.com"; Port = 443; Status = "Fail"; }

    $results += New-Object PSObject -Property @{Name = "Office 365 Portal"; Host = "outlook.office365.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Office 365 Portal"; Host = "home.office.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Office 365 Portal"; Host = "portal.office.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Office 365 Portal"; Host = "agent.office.net"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Office 365 Portal"; Host = "www.office.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Office 365 Portal"; Host = "portal.microsoftonline.com"; Port = 443; Status = "Fail"; }

    $results += New-Object PSObject -Property @{Name = "Shared infrastructure"; Host = "clientlog.portal.office.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Shared infrastructure"; Host = "nexus.officeapps.live.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Shared infrastructure"; Host = "nexusrules.officeapps.live.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Shared infrastructure"; Host = "accounts.office.net"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Shared infrastructure"; Host = "account.office.net"; Port = 443; Status = "Fail"; }

    $results += New-Object PSObject -Property @{Name = "Shared help and support"; Host = "support.office.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Shared help and support"; Host = "products.office.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Shared help and support"; Host = "templates.office.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Shared help and support"; Host = "contentstorage.osi.office.net"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Shared help and support"; Host = "technet.microsoft.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Shared help and support"; Host = "amp.azure.net"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Shared help and support"; Host = "assets.onestore.ms"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Shared help and support"; Host = "auth.gfx.ms"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Shared help and support"; Host = "browser.pip.aria.microsoft.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Shared help and support"; Host = "c.microsoft.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Shared help and support"; Host = "c1.microsoft.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Shared help and support"; Host = "connect.facebook.net"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Shared help and support"; Host = "dgps.support.microsoft.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Shared help and support"; Host = "mem.gfx.ms"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Shared help and support"; Host = "platform.linkedin.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Shared help and support"; Host = "support.content.office.net"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Shared help and support"; Host = "video.osi.office.net"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Shared help and support"; Host = "videocontent.osi.office.net"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Shared help and support"; Host = "videoplayercdn.osi.office.net"; Port = 443; Status = "Fail"; }

    $results += New-Object PSObject -Property @{Name = "Microsoft Azure RemoteApp"; Host = "dc.services.visualstudio.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Microsoft Azure RemoteApp"; Host = "liverdcxstorage.blob.core.windowsazure.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Microsoft Azure RemoteApp"; Host = "telemetry.remoteapp.windowsazure.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Microsoft Azure RemoteApp"; Host = "vortex.data.microsoft.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Microsoft Azure RemoteApp"; Host = "www.remoteapp.windowsazure.com"; Port = 443; Status = "Fail"; }

    $results += New-Object PSObject -Property @{Name = "Office 365 Management Pack for Operations Manager"; Host = "office365servicehealthcommunications.cloudapp.net"; Port = 443; Status = "Fail"; }

    $results += New-Object PSObject -Property @{Name = "Security and Compliance export"; Host = "protection.office.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Security and Compliance export"; Host = "office365zoom.cloudapp.net"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Security and Compliance export"; Host = "equivio.office.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Security and Compliance export"; Host = "compliance.outlook.com"; Port = 443; Status = "Fail"; }

    $results += New-Object PSObject -Property @{Name = "Office 365 Management APIs"; Host = "manage.office.com"; Port = 443; Status = "Fail"; }

    $results += New-Object PSObject -Property @{Name = "Graph API"; Host = "Graph.microsoft.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Graph API"; Host = "Graph.windows.net"; Port = 443; Status = "Fail"; }

    $results += New-Object PSObject -Property @{Name = "Discovery Service API"; Host = "api.office.com"; Port = 443; Status = "Fail"; }

    $results += New-Object PSObject -Property @{Name = "3rd party office integration."; Host = "firstpartyapps.oaspapps.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "3rd party office integration."; Host = "prod.firstpartyapps.oaspapps.com.akadns.net"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "3rd party office integration."; Host = "telemetryservice.firstpartyapps.oaspapps.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "3rd party office integration."; Host = "wus-firstpartyapps.oaspapps.com"; Port = 443; Status = "Fail"; }

    $results += New-Object PSObject -Property @{Name = "Microsoft Groups"; Host = "groupsapi-prod.outlookgroups.ms"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Microsoft Groups"; Host = "groupsapi2-prod.outlookgroups.ms"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Microsoft Groups"; Host = "groupsapi3-prod.outlookgroups.ms"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Microsoft Groups"; Host = "groupsapi4-prod.outlookgroups.ms"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Microsoft Groups"; Host = "sdk.hockeyapp.net"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Microsoft Groups"; Host = "rink.hockeyapp.net"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Microsoft Groups"; Host = "api.localytics.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Microsoft Groups"; Host = "analytics.localytics.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Microsoft Groups"; Host = "outlook.uservoice.com"; Port = 443; Status = "Fail"; }

    $results += New-Object PSObject -Property @{Name = "Authentication and identity"; Host = "api.login.microsoftonline.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Authentication and identity"; Host = "clientconfig.microsoftonline-p.net"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Authentication and identity"; Host = "device.login.microsoftonline.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Authentication and identity"; Host = "hip.microsoftonline-p.net"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Authentication and identity"; Host = "hipservice.microsoftonline.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Authentication and identity"; Host = "login.microsoft.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Authentication and identity"; Host = "login.microsoftonline.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Authentication and identity"; Host = "logincert.microsoftonline.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Authentication and identity"; Host = "loginex.microsoftonline.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Authentication and identity"; Host = "login-us.microsoftonline.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Authentication and identity"; Host = "login.microsoftonline-p.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Authentication and identity"; Host = "nexus.microsoftonline-p.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Authentication and identity"; Host = "stamp2.login.microsoftonline.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Authentication and identity"; Host = "login.windows.net"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Authentication and identity"; Host = "accesscontrol.windows.net"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Authentication and identity"; Host = "secure.aadcdn.microsoftonline-p.com"; Port = 443; Status = "Fail"; }

    $results += New-Object PSObject -Property @{Name = "Multi-factor authentication (MFA)"; Host = "account.activedirectory.windowsazure.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Multi-factor authentication (MFA)"; Host = "secure.aadcdn.microsoftonline-p.com"; Port = 443; Status = "Fail"; }

    $results += New-Object PSObject -Property @{Name = "DirSync"; Host = "login.windows.net"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "DirSync"; Host = "provisioningapi.microsoftonline.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "DirSync"; Host = "adminwebservice.microsoftonline.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "DirSync"; Host = "mscrl.microsoft.com"; Port = 443; Status = "Fail"; }

    $results += New-Object PSObject -Property @{Name = "Azure AD Connect"; Host = "login.windows.net"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Azure AD Connect"; Host = "provisioningapi.microsoftonline.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Azure AD Connect"; Host = "adminwebservice.microsoftonline.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Azure AD Connect"; Host = "mscrl.microsoft.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Azure AD Connect"; Host = "secure.aadcdn.microsoftonline-p.com"; Port = 443; Status = "Fail"; }

    $results += New-Object PSObject -Property @{Name = "Azure AD Connect Health"; Host = "management.azure.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Azure AD Connect Health"; Host = "policykeyservice.dc.ad.msft.net"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Azure AD Connect Health"; Host = "secure.aadcdn.microsoftonline-p.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Azure AD Connect Health"; Host = "login.windows.net"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Azure AD Connect Health"; Host = "login.microsoftonline.com"; Port = 443; Status = "Fail"; }

    $results += New-Object PSObject -Property @{Name = "Office 365 Management Pack for Operations Manager"; Host = "office365servicehealthcommunications.cloudapp.net"; Port = 443; Status = "Fail"; }

    $results += New-Object PSObject -Property @{Name = "Exchange Online"; Host = "smtp.office365.com"; Port = 587; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Exchange Online"; Host = "outlook.office365.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Exchange Online"; Host = "xsi.outlook.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Exchange Online"; Host = "r1.res.office365.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Exchange Online"; Host = "r3.res.office365.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Exchange Online"; Host = "r4.res.office365.com"; Port = 443; Status = "Fail"; }


    $results += New-Object PSObject -Property @{Name = "Skype for Business Online"; Host = "skypemaprdsitus.trafficmanager.net"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Skype for Business Online"; Host = "pipe.skype.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Skype for Business Online"; Host = "quicktips.skypeforbusiness.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Skype for Business Online"; Host = "swx.cdn.skype.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Skype for Business Online"; Host = "a.config.skype.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Skype for Business Online"; Host = "b.config.skype.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Skype for Business Online"; Host = "config.edge.skype.com"; Port = 443; Status = "Fail"; }


    $results += New-Object PSObject -Property @{Name = "SharePoint Online and OneDrive for Business"; Host = "cdn.sharepointonline.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "SharePoint Online and OneDrive for Business"; Host = "Static.sharepointonline.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "SharePoint Online and OneDrive for Business"; Host = "prod.msocdn.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "SharePoint Online and OneDrive for Business"; Host = "spoprod-a.akamaihd.net"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "SharePoint Online and OneDrive for Business"; Host = "publiccdn.sharepointonline.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "SharePoint Online and OneDrive for Business"; Host = "privatecdn.sharepointonline.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "SharePoint Online and OneDrive for Business"; Host = "oneclient.sfx.ms"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "SharePoint Online and OneDrive for Business"; Host = "https://officeclient.microsoft.com/config16"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "SharePoint Online and OneDrive for Business"; Host = "http://odc.officeapps.live.com/odc/emailhrd"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "SharePoint Online and OneDrive for Business"; Host = "login.microsoftonline.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "SharePoint Online and OneDrive for Business"; Host = "wns.windows.com"; Port = 443; Status = "Fail"; }



    $results += New-Object PSObject -Property @{Name = "Office 365 Video"; Host = "ajax.aspnetcdn.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Office 365 Video"; Host = "r3.res.outlook.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Office 365 Video"; Host = "spoprod-a.akamaihd.net"; Port = 443; Status = "Fail"; }

    $results += New-Object PSObject -Property @{Name = "Sway"; Host = "www.sway.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Sway"; Host = "eus-www.sway.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Sway"; Host = "eus-000.www.sway.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Sway"; Host = "eus-www.sway-cdn.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Sway"; Host = "wus-www.sway-cdn.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Sway"; Host = "eus-www.sway-extensions.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Sway"; Host = "wus-www.sway-cdn.com"; Port = 443; Status = "Fail"; }

    $results += New-Object PSObject -Property @{Name = "Planner"; Host = "tasks.office.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Planner"; Host = "controls.office.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Planner"; Host = "cus-000.tasks.osi.office.net"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Planner"; Host = "ea-000.tasks.osi.office.net"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Planner"; Host = "eus-zzz.tasks.osi.office.net"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Planner"; Host = "neu-000.tasks.osi.office.net"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Planner"; Host = "sea-000.tasks.osi.office.net"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Planner"; Host = "weu-000.tasks.osi.office.net"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Planner"; Host = "wus-000.tasks.osi.office.net"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Planner"; Host = "outlook.office365.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Planner"; Host = "www.outlook.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Planner"; Host = "clientlog.portal.office.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Planner"; Host = "ajax.aspnetcdn.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Planner"; Host = "prod.msocdn.com"; Port = 443; Status = "Fail"; }

    $results += New-Object PSObject -Property @{Name = "Microsoft Teams"; Host = "teams.skype.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Microsoft Teams"; Host = "teams.microsoft.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Microsoft Teams"; Host = "api.teams.skype.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Microsoft Teams"; Host = "img.teams.skype.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Microsoft Teams"; Host = "webhook.teams.skype.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Microsoft Teams"; Host = "statics.teams.skype.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Microsoft Teams"; Host = "statics.teams.microsoft.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Microsoft Teams"; Host = "bots.teams.skype.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Microsoft Teams"; Host = "settings.teams.skype.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Microsoft Teams"; Host = "emails.teams.skype.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Microsoft Teams"; Host = "emails.teams.skype.net"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Microsoft Teams"; Host = "prod.registrar.skype.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Microsoft Teams"; Host = "prod.tpc.skype.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Microsoft Teams"; Host = "amer-client-ss.msg.skype.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Microsoft Teams"; Host = "amer-server-ss.msg.skype.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Microsoft Teams"; Host = "us-api.asm.skype.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Microsoft Teams"; Host = "emea-client-ss.msg.skype.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Microsoft Teams"; Host = "emea-server-ss.msg.skype.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Microsoft Teams"; Host = "eu-api.asm.skype.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Microsoft Teams"; Host = "apac-client-ss.msg.skype.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Microsoft Teams"; Host = "apac-server-ss.msg.skype.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Microsoft Teams"; Host = "ea-api.asm.skype.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Microsoft Teams"; Host = "mobile.pipe.aria.microsoft.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Microsoft Teams"; Host = "ssdesktopbuild.blob.core.windows.net"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Microsoft Teams"; Host = "s-0001.s-msedge.net"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Microsoft Teams"; Host = "s-0002.s-msedge.net"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Microsoft Teams"; Host = "s-0004.s-msedge.net"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Microsoft Teams"; Host = "scsquery-ss-us.trafficmanager.net"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Microsoft Teams"; Host = "scsquery-ss-eu.trafficmanager.net"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Microsoft Teams"; Host = "scsquery-ss-asia.trafficmanager.net"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Microsoft Teams"; Host = "a.config.skype.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Microsoft Teams"; Host = "b.config.skype.com"; Port = 443; Status = "Fail"; }




    $results += New-Object PSObject -Property @{Name = "Office 365 remote analyzer tools"; Host = "testconnectivity.microsoft.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Office 365 remote analyzer tools"; Host = "client.hip.live.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Office 365 remote analyzer tools"; Host = "wu.client.hip.live.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Office 365 remote analyzer tools"; Host = "support.microsoft.com"; Port = 443; Status = "Fail"; }


    $results += New-Object PSObject -Property @{Name = "OneNote notebooks"; Host = "www.onenote.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "OneNote notebooks"; Host = "cdn.onenote.net"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "OneNote notebooks"; Host = "cdn.optimizely.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "OneNote notebooks"; Host = "Ajax.aspnetcdn.com"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "OneNote notebooks"; Host = "apis.live.net"; Port = 443; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "OneNote notebooks"; Host = "www.onedrive.com"; Port = 443; Status = "Fail"; }










    foreach ($result in $results) {
        $result | Add-Member MemberSet PSStandardMembers $PSStandardMembers
    }

   
      
    foreach ($result in $results) {
    
        $url = $result | select -ExpandProperty Host
        $port = $result | select -ExpandProperty Port


        $status = Test-NetConnection -ComputerName $url -Port $port -WarningAction SilentlyContinue | select -ExpandProperty TCPTestSucceeded

        if($status)
        {
            $result.Status = 'Pass'
        }
    
    

    }


    return $results;
}

}

Check-OfficeProPlusPorts