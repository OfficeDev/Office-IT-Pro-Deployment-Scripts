#Get New Office Users

Two functions to identify licensed Office 365 users and track the dates they were enabled or disabled.

###Pre-requisites

The script requires the Azure Active Directory Services Module, previously known as the Microsoft Online Services Module.

The module information can be found at https://technet.microsoft.com/en-us/library/jj151815.aspx.

###Update-UserLicenseData

Finds all of the MSOLUsers that are licensed with the specified plan and stores them in a CSV.
An extra field is populated in the CSV that specifies the earliest point at which the CSV
knew the user was licensed (LicensedAsOf field). If a user doesn't show up in the licensed list at a later date
the function takes note and populates another field in the CSV with that date (DelicensedAsOf).

####Examples

1. Open a PowerShell console.

		From the Run dialog type PowerShell.
	
2. Change directory to the location where the PowerShell Script is saved.

		Example: cd C:\PowerShellScripts
	
3. Run the Update-UserLicenseData.ps1 script.

		Type . .\Get-NewOfficeUsers.ps1

		By including the additional period before the relative script path you are 'Dot-Sourcing' 
		the PowerShell function in the script into your PowerShell session which will allow you to 
		run the function from the console.
		
4. Run the Update-UserLicenseData.ps1 script and you will be prompted for the Office 365 username and password

		Type Update-UserLicenseData -Credentials (Get-Credential)
		
###Get-RecentlyLicensedUsers

Get a list of users that were licensed after the specified date according to the specified CSV.
It is important to run the Update-UserLicenseData.ps1 prior to using this script.

####Examples

1. Open a PowerShell console.

		From the Run dialog type PowerShell.
	
2. Change directory to the location where the PowerShell Script is saved.

		Example: cd C:\PowerShellScripts
	
3. Run the Get-NewOfficeUsers.ps1 script. By default this will show you the users that have been created in the last week.

		Type . .\Get-NewOfficeUsers.ps1

###Recommended Use Case

Proper use of Update and Email scripts should involve running them as a scheduled task using a service 
account (not personal) because the password will be put in plain text in the scheduled task.

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
		-or-
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
