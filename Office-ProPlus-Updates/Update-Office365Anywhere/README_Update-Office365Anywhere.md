##**Update Office 365 Anywhere**

This function is designed to provide way for Office Click-To-Run clients to have the ability to update themselves from a managed network source
or from the Internet depending on the availability of the primary update source.  The idea behind this is if users have laptops and are mobile 
they may not recieve updates if they are not able to be in the office on regular basis.  This functionality is available with this function but it's 
use can be controller by the parameter -EnableUpdateAnywhere.  This function also provides a way to initiate an update and the script will wait
for the update to complete before exiting. Natively starting an update executable does not wait for the process to complete before exiting and
in certain scenarios it may be useful to have the update process wait for the update to complete.

###**Running the script**

1. dsfsdf

		From the Run dialog type PowerShell.

3. Change directory to the location where the PowerShell Script is saved. This directory must contain the files that are in the *Setup-SCCMOfficeUpdates* folder.

		Example: cd C:\PowerShellScripts

4. Type the following in the elevated PowerShell Session

		 . .\Setup-SCCMOfficeUpdates.ps1
         

