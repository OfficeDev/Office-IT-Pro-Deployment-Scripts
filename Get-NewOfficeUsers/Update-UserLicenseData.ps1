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

.PARAMETER Username
Used to generate Credentials for connecting to MSOL service

.PARAMETER Password
Used to generate Credentials for connecting to MSOL service

.Example
.\Update-UserLicenseData.ps1 -ServiceName "OFFICESUBSCRIPTION" -CSVPath "$env:Public\Documents\LicensedUsers.csv"
Get list of Users that are licensed for OFFICESUBSCRIPTION service plan and store the results in public documents.
The user will be prompted for their credentials.

.Example
.\Update-UserLicenseData.ps1 -ServiceName "OFFICESUBSCRIPTION" -CSVPath "$env:Public\Documents\LicensedUsers.csv" -Credentials $creds
Get list of Users that are licensed for OFFICESUBSCRIPTION service plan and store the results in public documents.
The user won't be prompted for their credentials

.Notes
Proper use of this script should involve running this as a scheduled task

    1.  On the system that the task will be run from, open the Windows Task Scheduler. 
        This can be found in the Start menu, under Start > Administrative Tools.

    2.  In the Task Scheduler, select the Create Task option under the Actions heading 
        on the right-hand side.

    3.  Enter a name for the task, and give it a description (the description is optional 
        and not required).

    4.  In the General tab, go to the Security options heading and specify the user account 
        that the task should be run under. Change the settings so the task will run if the 
        user is logged in or not.

    5.  Next, select the Triggers tab, and click New to add a new trigger for the scheduled 
        task. This new task should use the On a schedule option. The start date can be set 
        to a desired time, and the frequency and duration of the task can be set based on 
        your specific needs. Click OK when your desired settings are entered.

    6.  Next, go to the Actions tab and click New to set the action for this task to run. 
        Set the Action to Start a program.

    7.  In the Program/script box enter "PowerShell."
        In the Add arguments (optional) box enter the value:

         .\Update-UserLicenseData.ps1 -ServiceName [ServiceName] -CSVPath [Full\file\path\of.csv] -Username [username] -Password [password])

    8.  Then, in the Start in (optional) box, add the location of the folder that contains 
        your PowerShell script.

        Note: The location used in the Start in box will also be used for storing the scheduled task 
        run times, the job history for the copies, and any additional logging that may occur.
        Click OK when all the desired settings are made.

    9. Next, set any other desired settings in the Conditions and Settings tabs. You can also set up 
        additional actions, such as emailing an Administrator each time the script is run.

    10. Once all the desired actions have been made (or added), click OK. The task will be immediately 
        set, and is ready to run.

        The scheduling of this task is complete, and is now ready to run based on the entered settings.

#>
[CmdletBinding(DefaultParameterSetName="PSCredential")]
Param(

    [Parameter()]
    [string] $ServiceName,

    [Parameter()]
    [string] $CSVPath,

    [Parameter(ParameterSetName="PSCredential")]
    [PSCredential] $Credentials,

    [Parameter(ParameterSetName="UsernamePassword")]
    [string] $Username,

    [Parameter(ParameterSetName="UsernamePassword")]
    [string] $Password

)

Process{
    if($PSCmdlet.ParameterSetName -eq "UsernamePassword"){
        $PWord = ConvertTo-SecureString –String $Password –AsPlainText -Force
        $Credentials = New-Object –TypeName System.Management.Automation.PSCredential –ArgumentList $Username, $PWord
    }else{
        $Credentials = (Get-Credential)
    }

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
        

        Foreach($ImportedUser in $ImportedCSV){
            $CheckUser = $LicensedUsers | ? ObjectId -eq $ImportedUser.ObjectId
            if($CheckUser -eq $Null){
                $ImportedUser.DelicensedAsOf = "$(Get-Date -Format "yyyy-MM-dd hh:mm")"
            }
        }

        Foreach($LicensedUser in $LicensedUsers){
            $CheckUser = $ImportedCSV | ? ObjectId -eq $LicensedUser.ObjectId
            if($CheckUser -eq $Null){
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
