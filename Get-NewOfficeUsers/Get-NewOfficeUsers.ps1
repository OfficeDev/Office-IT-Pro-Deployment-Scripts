
function Update-UserLicenseData{
    <#
    .SYNOPSIS
    Finds all the MSOLUsers that are licensed with the specified plan and stores them in a csv

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
    Update-UserLicenseData -ServiceName "OFFICESUBSCRIPTION" -CSVPath "$env:Public\Documents\LicensedUsers.csv"
    Get list of Users that are licensed for OFFICESUBSCRIPTION service plan and store the results in public documents.
    The user will be prompted for their credentials.

    .Example
    Update-UserLicenseData -ServiceName "OFFICESUBSCRIPTION" -CSVPath "$env:Public\Documents\LicensedUsers.csv" -Credentials $creds
    Get list of Users that are licensed for OFFICESUBSCRIPTION service plan and store the results in public documents.
    The user won't be prompted for their credentials

    #>
    Param(

        [Parameter()]
        [string] $ServiceName,

        [Parameter()]
        [string] $CSVPath,

        [Parameter()]
        [PSCredential] $Credentials = (Get-Credential)

    )

    Process{

        Connect-MsolService -Credential $Credentials
        $LicensedUsers = Get-MsolUser | ? IsLicensed -eq $True | Select DisplayName, Licenses, LiveId, ObjectId, SignInName 
        $O365Users = new-object PSObject[] 1;
        foreach($User in $LicensedUsers){
            :LicenseLoop foreach($License in $User.Licenses){
                foreach($ServiceStatus in $License.ServiceStatus){
                    if($ServiceStatus.ServicePlan.ServiceName -eq $ServiceName){
                        $O365Users += $User
                        break LicenseLoop
                    }
                }
            }
        }
        
        foreach($User in $O365Users){
            if($User -ne $Null){
                Add-Member -InputObject $User -MemberType NoteProperty -Name LicensedAsOf -Value "$(Get-Date -Format "yyyy-MM-dd hh:mm")"
                Add-Member -InputObject $User -MemberType NoteProperty -Name DelicensedAsOf -Value "-"
            }
        }

        if(Test-Path $CSVPath){
            $ImportedCSV = Import-Csv $CSVPath
        

            Foreach($importedUser in $ImportedCSV){
                $test123 = $O365Users | ? ObjectId -eq $importedUser.ObjectId
                if($test123 -eq $null){
                    $importedUser.DelicensedAsOf = "$(Get-Date -Format "yyyy-MM-dd hh:mm")"
                }
            }

            Foreach($O365User in $O365Users){
                $test123 = $ImportedCSV | ? ObjectId -eq $O365User.ObjectId
                if($test123 -eq $Null){
                    if($O365User -ne $Null){
                        $ImportedCSV += $O365User
                    }
                }
            }
            $ImportedCSV | Export-Csv $CSVPath -NoTypeInformation
        }else{
            $O365Users | ? ObjectId -ne $Null | Export-Csv $CSVPath -NoTypeInformation
        }
    }
}

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