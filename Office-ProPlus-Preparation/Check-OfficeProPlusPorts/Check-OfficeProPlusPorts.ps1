Function Check-OfficeProPlusPorts {
<#
.Synopsis

.DESCRIPTION

.NOTES   
Name: Check-OfficeProPlusPorts
Version: 1.0.0
DateCreated: 2015-11-10
DateUpdated: 2015-11-10

.LINK
https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts

.EXAMPLE
Get-OfficeVersion

Description:
Will return the locally installed Office product

.EXAMPLE
Check-OfficeProPlusPorts

Description:

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
    $results += New-Object PSObject -Property @{Name = "Installation/Update Contet"; Host = "officecdn.microsoft.com"; Port = 80; Status = "Fail"; }
    $results += New-Object PSObject -Property @{Name = "Online Help Services"; Host = "go.microsoft.com"; Port = 80; Status = "Fail"; }


    foreach ($result in $results) {
      $result | Add-Member MemberSet PSStandardMembers $PSStandardMembers
    }

    
    foreach ($result in $results) {
    
    $url = $result | select -ExpandProperty Host
    $port = $result | select -ExpandProperty Port

    $status = Test-NetConnection -ComputerName $url -Port $port

    Write-Output $status


    }


    #return $results;
}

}

Check-OfficeProPlusPorts