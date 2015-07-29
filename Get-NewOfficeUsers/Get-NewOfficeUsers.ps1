
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
        $Users = Get-MsolUser | ? IsLicensed -eq $True | Select DisplayName, Licenses, LiveId, ObjectId, SignInName 

        #Get list of users with the correct service plan
        $LicensedUsers = new-object PSObject[] 1;
        foreach($User in $Users){
            :LicenseLoop foreach($License in $User.Licenses){
                foreach($ServiceStatus in $License.ServiceStatus){
                    if($ServiceStatus.ServicePlan.ServiceName -eq $ServiceName){
                        $LicensedUsers += $User
                        break LicenseLoop
                    }
                }
            }
        }
        
        #Add tracking properties
        foreach($User in $LicensedUsers){
            if($User -ne $Null){
                Add-Member -InputObject $User -MemberType NoteProperty -Name LicensedAsOf -Value "$(Get-Date -Format "yyyy-MM-dd hh:mm")"
                Add-Member -InputObject $User -MemberType NoteProperty -Name DelicensedAsOf -Value "-"
            }
        }

        #Check if CSV exists
        if(Test-Path $CSVPath){
            #if CSV exists, import it and compare and update values
            $ImportedCSV = Import-Csv $CSVPath
        

            Foreach($importedUser in $ImportedCSV){
                $test123 = $LicensedUsers | ? ObjectId -eq $importedUser.ObjectId
                if($test123 -eq $null){
                    $importedUser.DelicensedAsOf = "$(Get-Date -Format "yyyy-MM-dd hh:mm")"
                }
            }

            Foreach($LicensedUser in $LicensedUsers){
                $test123 = $ImportedCSV | ? ObjectId -eq $LicensedUser.ObjectId
                if($test123 -eq $Null){
                    if($LicensedUser -ne $Null){
                        $ImportedCSV += $LicensedUser
                    }
                }
            }
            $ImportedCSV | Export-Csv $CSVPath -NoTypeInformation
        }else{
            #If csv does not exist, export data
            $LicensedUsers | ? ObjectId -ne $Null | Export-Csv $CSVPath -NoTypeInformation
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