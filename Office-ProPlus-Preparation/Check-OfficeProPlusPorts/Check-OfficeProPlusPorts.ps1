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

    $results += New-Object PSObject -Property @{Name = "Renew Product Key"; Host = "activation.sls.microsoft.com"; Port = 443; Status = ""; }
    $results += New-Object PSObject -Property @{Name = "Validate Certificates"; Host = "crl.microsoft.com"; Port = 443; Status = ""; }
    $results += New-Object PSObject -Property @{Name = "Validate Certificates"; Host = "crl.microsoft.com"; Port = 80; Status = ""; }

    foreach ($result in $results) {






    }

    foreach ($result in $results) {
      $result | Add-Member MemberSet PSStandardMembers $PSStandardMembers
    }

    return $results;
}

}

Check-OfficeProPlusPorts