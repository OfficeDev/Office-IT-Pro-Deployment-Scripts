
function Update-UserLicenseData{
    
    Param(

        [Parameter()]
        [string] $ServiceName,

        [Parameter()]
        [string] $CSVPath

    )

    Process{
        $creds = Get-Credential
        Connect-MsolService -Credential $creds
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

        
        $O365Users | %{ Add-Member -InputObject $_ -MemberType NoteProperty -Name LicensedAsOf -Value "$(Get-Date -Format "yyyy-MM-dd")"}

        if(Test-Path $CSVPath){
            $importedCSV = Import-Csv $CSVPath
        

            Foreach($importedUser in $importedCSV){
                $test123 = $O365Users | ? ObjectId -eq $importedUser.ObjectId
                if($test123 -eq $null){
                    $importedCSV.Remove($importedUser);
                }
            }

            Foreach($O365User in $O365Users){
                $test123 = $importedCSV | ? ObjectId -eq $O365User.ObjectId
                if($test123 -eq $null){
                    $importedCSV += $O365User
                }
            }
            $importedCSV | Export-Csv $CSVPath
        }else{
            $O365Users | Export-Csv $CSVPath
        }
    }
}

function Get-RecentlyLicensedUsers{

    $NewUsers = new-object PSObject[] 1;

    $importedCSV = Import-Csv $CSVPath

    [DateTime]$CutOffDate

    foreach($user in $importedCSV){
        if($CutOffDate.CompareTo((Get-Date($user.LicensedAsOf))) -lt 0){
            $NewUsers += $user
        }
    }

    return $NewUsers

}