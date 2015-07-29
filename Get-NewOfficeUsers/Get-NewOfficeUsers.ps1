function Get-RecentlyLicensedUsers{
    <#
    .SYNOPSIS
    Get a list of users the were licensed after the specified date according to the specified csv

    .DESCRIPTION
    Finds all the MSOLUsers that are licensed with the specified plan and stores them in a csv.
    It populates an extra field in the csv the specifies the earliest point at which the csv
    knew the user was licensed (LicensedAsOf field). If a user doesn't show up in the licensed list at a later date
    the function takes note and populates another field in the csv with that date (DelicensedAsOf)

    .PARAMETER ServiceName
    The Name of the Service Plan that you wish to track licensed users for

    .PARAMETER CSVPath
    The Full file path of the CSV you wish the data to be tracked in

    .PARAMETER Credentials
    Credentials for connecting to MSOL service

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

        foreach($user in $importedCSV){
            if($CutOffDate.CompareTo((Get-Date($user.LicensedAsOf))) -lt 0){
                $NewUsers += $user
            }
        }

        return $NewUsers
    }

}