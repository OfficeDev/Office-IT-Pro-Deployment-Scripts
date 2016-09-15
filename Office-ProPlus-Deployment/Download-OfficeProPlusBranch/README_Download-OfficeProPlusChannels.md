### Download Office ProPlus Channels
This PowerShell function download the latest Office ProPlus branch to one location

[README](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/wiki/Readme_Download_OfficeProPlusBranch)


1. Open a PowerShell console.

		From the Run dialog type PowerShell 

2. Change directory to the location where the PowerShell Script is saved.

		Example: cd C:\PowerShellScripts

3. Dot-Source the Download-OfficeProPlusChannels.ps1 functions into your current session.

		Type . .\Download-OfficeProPlusChannels.ps1

		
		By including the additional period before the relative script path you are 'Dot-Sourcing' 
		the PowerShell function in the script into your PowerShell session which will allow you to 
		run the function from the console.

4. Run the Download-OfficeProPlusChannels function to download the Office 365 installer files for the channel specified. 

		Example: Download-OfficeProPlusChannels -Version %16.0.xxxx.xxxx% -Languages %en-us% -Bitness v32 -Channels current 