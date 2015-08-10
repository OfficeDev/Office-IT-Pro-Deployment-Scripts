$90DayBody =  
"Greetings {0},

    90Body

Best Regards,
SignOff Name"

$90DaySubject = "90 Days"

$30DayBody =  
"Greetings {0},

    30Body

Best Regards,
SignOff Name"

$30DaySubject = "30 Days"

$5DayBody =  
"Greetings {0},

    5Body

Best Regards,
SignOff Name"

$5DaySubject = "5 Days"

$1DayBody = 
"Greetings {0},

    1Body

Best Regards,
SignOff Name"

$1DaySubject = "1 Days"



Function Update-UserLicenseData {

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
.\Update-UserLicenseData.ps1
Get list of Users that are licensed for OFFICESUBSCRIPTION service plan and store the results in an AppData Folder
The user will be prompted for their credentials.

.Example
.\Update-UserLicenseData.ps1 -ServiceName "OFFICESUBSCRIPTION"
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
        (For Windows 7 and upward, you can just search for Task Scheduler in the start
        menu)

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

    7.  In the Program/script box enter "PowerShell.exe"
        In the Add arguments (optional) box enter the value:

        . ./Get-NewOfficeUsers.ps1;Update-UserLicenseData -ServiceName [ServiceName] -Username [username] -Password [password])

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
    [string] $ServiceName = "OFFICESUBSCRIPTION",

    [Parameter()]
    [string] $CSVPath = "$env:APPDATA\Microsoft\OfficeAutomation\OfficeLicenseTracking.csv",

    [Parameter(ParameterSetName="PSCredential")]
    [PSCredential] $Credentials,

    [Parameter(ParameterSetName="UsernamePassword")]
    [string] $Username,

    [Parameter(ParameterSetName="UsernamePassword")]
    [string] $Password

)

Process{
    if($PSCmdlet.ParameterSetName -eq "UsernamePassword")
    {
        $PWord = ConvertTo-SecureString –String $Password –AsPlainText -Force
        $Credentials = New-Object –TypeName System.Management.Automation.PSCredential –ArgumentList $Username, $PWord
    } else {
      if (!($Credentials)) {
        $Credentials = (Get-Credential)
      }
    }

    if ($Credentials) {
        $Domain = $Credentials.UserName.Split('@')[1]
        $CSVPath = "$env:APPDATA\Microsoft\OfficeAutomation\OfficeLicenseTracking-$Domain.csv"
    
        Write-host
        Write-host "Connecting to Office 365..."

        Connect-MsolService -Credential $Credentials

        Write-host "Retrieving User List..."

        $Users = Get-MsolUser | ? IsLicensed -eq $True | Select DisplayName, Licenses, LiveId, ObjectId, SignInName, UserPrincipalName 
    
        #Get list of users with the correct service plan
        $LicensedUsers = new-object PSObject[] 1;
        foreach($User in $Users){
            :LicenseLoop foreach($License in $User.Licenses){
                foreach($ServiceStatus in $License.ServiceStatus){
                    if($ServiceStatus.ServicePlan.ServiceName -eq $ServiceName){
                        $LicensedUsers += $User
                        break LicenseLoop
                    } #End if
                } #End ServiceStatus Foreach
            } #End License Foreach
        } #End User in Users Foreach
        
        #Add tracking properties
        foreach($User in $LicensedUsers){
            if($User -ne $Null){
                Add-Member -InputObject $User -MemberType NoteProperty -Name LicensedAsOf -Value "$(Get-Date -Format "yyyy-MM-dd hh:mm")"
                Add-Member -InputObject $User -MemberType NoteProperty -Name DelicensedAsOf -Value "-"
                Add-Member -InputObject $User -MemberType NoteProperty -Name Day1EmailSent -Value "$false"
                Add-Member -InputObject $User -MemberType NoteProperty -Name Day5EmailSent -Value "$false"
                Add-Member -InputObject $User -MemberType NoteProperty -Name Day30EmailSent -Value "$false"
                Add-Member -InputObject $User -MemberType NoteProperty -Name Day90EmailSent -Value "$false"
            } #End if
        } #End User in LicensedUsers Foreach

        $pathSplit = Split-Path -Path $CSVPath
        $createDir = [system.io.directory]::CreateDirectory($pathSplit)

        #Check if CSV exists
        if(Test-Path $CSVPath){
            #if CSV exists, import it and compare and update values
            $ImportedCSV = Import-Csv $CSVPath
        

            Foreach($ImportedUser in $ImportedCSV){
                $CheckUser = $LicensedUsers | ? ObjectId -eq $ImportedUser.ObjectId
                if($CheckUser -eq $Null){
                    $ImportedUser.DelicensedAsOf = "$(Get-Date -Format "yyyy-MM-dd hh:mm")"
                } #End if
            } #End ImportedUser Foreach

            Foreach($LicensedUser in $LicensedUsers){
                $CheckUser = $ImportedCSV | ? ObjectId -eq $LicensedUser.ObjectId
                if($CheckUser -eq $Null){
                    if($LicensedUser -ne $Null){
                        $ImportedCSV += $LicensedUser
                    } #end LicensedUser if
                } #end CheckUser if
            } #End LicensedUser Foreach
            $ImportedCSV | Export-Csv $CSVPath -NoTypeInformation

            Write-host "CSV File Updated: $CSVPath"
        } #End TestPath if
        else{
            #If csv does not exist, export data
            $LicensedUsers | ? ObjectId -ne $Null | Export-Csv $CSVPath -NoTypeInformation

            Write-host "CSV File Created: $CSVPath"
        } #End TestPath Else
    }#end Credentials If
}#End Process

}

Function Get-RecentlyLicensedUsers {

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
Get-RecentlyLicensedUsers 
Get list of Users that are were licensed in the last 7 days if the Update-UserLicenseData cmdlet has already been run

.Example
Get-RecentlyLicensedUsers -CutOffDate (Get-Date "2015-7-13) -CSVPath "$env:Public\Documents\LicensedUsers.csv"
Get list of Users that are were licensed after July 7, 2015 according to specified csv

#>

[CmdletBinding()]
Param(

    [Parameter()]
    [string] $CSVPath,

    [Parameter()]
    [DateTime] $CutOffDate

)

Begin {
    $defaultDisplaySet = 'DisplayName', 'SignInName'

    $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet(‘DefaultDisplayPropertySet’,[string[]]$defaultDisplaySet)
    $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
}#End Begin

Process{
    
    if (!($CutOffDate)) {
       $CutOffDate = (Get-Date).AddDays(-7)
    }#end !CutOffDate

    [System.IO.FileSystemInfo[]]$filePaths = @()

    if (!($CSVPath)) {
        $childItems = Get-ChildItem -Path "$env:APPDATA\Microsoft\OfficeAutomation" | where {$_.Extension.ToLower() -eq ".csv" }
        foreach ($csvFile in $childItems) {
           $filePaths += $csvFile
        }#end foreach csvFile
    } else {
       $filePaths += ([System.IO.FileInfo]"$CSVPath")
    }#end if/else !csvPath
    
    if ($filePaths.Length -eq 0) {
      Write-Host "No CSV File Exits. Please run Update-UserLicenseData to generate the CSV File."
    } #End if filePaths

    foreach ($csvFile in $filePaths) {
        $fileName = $csvFile.Name.Replace($csvFile.Extension, "")
        $domain = ""
        if ($fileName.Contains("-")) {
           $domain = $fileName.Split('-')[1]
        }#End if fileName
       
        if (($PSCmdlet.MyInvocation.PipelineLength -eq 1) -or `
            ($PSCmdlet.MyInvocation.PipelineLength -eq $PSCmdlet.MyInvocation.PipelinePosition)) {

            Write-Host ""
            Write-Host "Retrieving New Users Since: $CutOffDate"
            Write-Host ""
            if ($domain) {
               Write-Host "Domain: $domain"
            }#end If domain
        }#End if pipeline

        $NewUsers = new-object PSObject[] 1;

        $ImportedCSV = Import-Csv -LiteralPath $csvFile.FullName

        foreach($User in $ImportedCSV){
            if ($CutOffDate -lt (Get-Date($user.LicensedAsOf))) {
              if ($domain) {
                Add-Member -InputObject $User -MemberType NoteProperty -Name "Domain" -Value $domain
              }#End If domain
              $User | Add-Member MemberSet PSStandardMembers $PSStandardMembers

              $NewUsers += $User
            }#End if CutOffDate
        }#End Foreach User

        return $NewUsers
    }#End foreach csvFile
}#End Process

}

Function Send-RecentUserEmails{
<#
.SYNOPSIS
Sends emails to Users at intervals to encourage use of Office

.DESCRIPTION
Uses the data stored in the csv populated in Update-UserLicenseData.
Sends emails at 1, 5, 30, 90 days to users after the LicensedAsOf date.

.PARAMETER SmtpServer
The SmtpServer of the Email that will be used to send the emails

.PARAMETER CSVPath
The Full file path of the CSV where the data is tracked

.PARAMETER Credentials
Credentials for connecting to MSOL service

.PARAMETER Username
Used to generate Credentials for connecting to MSOL service

.PARAMETER Password
Used to generate Credentials for connecting to MSOL service

.Example
Send-RecentUserEmails
Sends emails to all the users in the proper intervals that have not received the emails yet.
Will look for the csv in the default location using the domain name from the credentials.
The user will be prompted for their credentials.

.Example
Send-RecentUserEmails -Credentials $creds
Sends emails to all the users in the proper intervals that have not received the emails yet.
Will look for the csv in the default location using the domain name from the credentials.
The user won't be prompted for their credentials.

.Notes
Proper use of this script should involve running this as a scheduled task

    1.  On the system that the task will be run from, open the Windows Task Scheduler. 
        This can be found in the Start menu, under Start > Administrative Tools.
        (For Windows 7 and upward, you can just search for Task Scheduler in the start
        menu)

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

    7.  In the Program/script box enter "PowerShell.exe"
        In the Add arguments (optional) box enter the value:

        . ./Get-NewOfficeUsers.ps1;Send-RecentUserEmails -SmtpServer [SmtpServer] -Username [username] -Password [password])

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
[CmdletBinding()]
Param(

    [Parameter()]
    [string] $SmtpServer = "smtp.office365.com",

    [Parameter()]
    [string] $CSVPath = "$env:APPDATA\Microsoft\OfficeAutomation\OfficeLicenseTracking.csv",

    [Parameter(ParameterSetName="PSCredential")]
    [PSCredential] $Credentials,

    [Parameter(ParameterSetName="UsernamePassword")]
    [string] $Username,

    [Parameter(ParameterSetName="UsernamePassword")]
    [string] $Password

)

Process{

    if($PSCmdlet.ParameterSetName -eq "UsernamePassword")
    {
        $PWord = ConvertTo-SecureString –String $Password –AsPlainText -Force
        $Credentials = New-Object –TypeName System.Management.Automation.PSCredential –ArgumentList $Username, $PWord
    } else {
      if (!($Credentials)) {
        $Credentials = (Get-Credential)
      }
    }

    if($Credentials){

        $Domain = $Credentials.UserName.Split('@')[1]
        $CSVPath = "$env:APPDATA\Microsoft\OfficeAutomation\OfficeLicenseTracking-$Domain.csv"
        $5DaysAgo = (Get-Date).AddDays(-5);
        $30DaysAgo = (Get-Date).AddDays(-30);
        $90DaysAgo = (Get-Date).AddDays(-90);

        $Users = Import-Csv $CSVPath

        $Denominator = $Users.Count;
        [int]$Progress = 0;
        Write-Progress -Activity "Sending Emails" -Status "Processing Users" -PercentComplete ($Progress/$Denominator)

        foreach( $User in $Users ){
        
            if((Get-Date($User.LicensedAsOf)) -gt $5DaysAgo -and $User.Day1EmailSent -eq $false){

                $Body = $1DayBody -f $User.DisplayName
                Send-MailMessage -To $User.UserPrincipalName -From $Credentials.UserName -Subject $1DaySubject -Body $Body -SmtpServer $SmtpServer -Credential $Credentials -UseSsl -Port "587"
                $User.Day1EmailSent = $true;

            }elseif((Get-Date($User.LicensedAsOf)) -gt $30DaysAgo -and $User.Day5EmailSent -eq $false){

                $Body = $5DayBody -f $User.DisplayName
                Send-MailMessage -To $User.UserPrincipalName -From $Credentials.UserName -Subject $5DaySubject -Body $Body -SmtpServer $SmtpServer -Credential $Credentials -UseSsl -Port "587"
                $User.Day5EmailSent = $true;

            }elseif((Get-Date($User.LicensedAsOf)) -gt $90DaysAgo -and $User.Day30EmailSent -eq $false){

                $Body = $30DayBody -f $User.DisplayName
                Send-MailMessage -To $User.UserPrincipalName -From $Credentials.UserName -Subject $30DaySubject -Body $Body -SmtpServer $SmtpServer -Credential $Credentials -UseSsl -Port "587"
                $User.Day30EmailSent = $true;

            }elseif($User.Day90EmailSent -eq $false){

                $Body = $90DayBody -f $User.DisplayName
                Send-MailMessage -To $User.UserPrincipalName -From $Credentials.UserName -Subject $90DaySubject -Body $Body -SmtpServer $SmtpServer -Credential $Credentials -UseSsl -Port "587"
                $User.Day90EmailSent = $true;

            }#End Date if/elses

            $Progress += 1
            Write-Progress -Activity "Sending Emails" -Status "Processing Users" -PercentComplete ($Progress/$Denominator)

        }#end Foreach User

        $Users | Export-Csv $CSVPath

    }#end if Credentials
}#End Process

}