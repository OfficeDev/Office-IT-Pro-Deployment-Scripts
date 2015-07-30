<#
.SYNOPSIS
Get a list of users the were licensed after the specified date according to the specified csv

.DESCRIPTION
Get a list of users the were licensed after the specified date according to the specified csv.
It is important to have run the Update-UserLicenseData.ps1 prior to using this script.

.PARAMETER CutOffDate
The cutoff date for how new you wish the return list of users to be

.PARAMETER CSVPath
The Full file path of the CSV with the data (should be the same as the path used for
Update-UserLicenseData.ps1).

.Example
Get-RecentlyLicensedUsers -CutOffDate (Get-Date "2015-7-13) -CSVPath "$env:Public\Documents\LicensedUsers.csv"
Get list of Users that are were licensed after July 7, 2015 according to specified csv

#>
Param(

    [Parameter()]
    [string] $CSVPath,

    [Parameter()]
    [DateTime] $CutOffDate

)

Process{

    $NewUsers = new-object PSObject[] 1;

    $ImportedCSV = Import-Csv $CSVPath

    foreach($User in $ImportedCSV){
        if($CutOffDate.CompareTo((Get-Date($user.LicensedAsOf))) -lt 0){
            $NewUsers += $User
        }
    }

    return $NewUsers
}