Param
(
	[Parameter(Mandatory=$True)]
	[String]$version

	[Parameter(Mandatory=$True)]
	[String]$path,

	[Parameter()]
	[String]$bitness = '64',

	[Parameter(Mandatory=$True)]
	[String]$siteId,

	[Parameter(Mandatory=$True)]
	[String]$username,
	
	[Parameter(Mandatory=$True)]
	[String]$plainTextPassword
	
	[Parameter(Mandatory=$True)]
	[String]$UpdateSourceConfigFileName = 'Configuration_UpdateSource.xml'

	[Parameter(Mandatory=$True)]
	[String]$UpdateTestGroupConfigFileName = 'Configuration_UpdateTestGroup.xml'
)

Begin
{
	
}

Process
{
	# Get Credentials required to connect to the Share Path
	$secpasswd = ConvertTo-SecureString $plainTextPassword -AsPlainText -Force
	$credentials = New-Object System.Management.Automation.PSCredential ($username, $secpasswd) # Alternative - $credentials = Get-Credential



	$c2rFileName = 'setup.exe'

	#1 - Set the correct version number to update Source location
	$sourceFilePath = "$path\$UpdateSourceConfigFileName"
	$sourceContent = [Xml] (Get-Content $sourceFilePath)
	$addNode = $sourceContent.Configuration.Add
	$addNode.OfficeClientEdition = $bitness
	$addNode.Version = $version

	$sourceContent.Save($sourceFilePath)

	$testGroupFilePath = "$path\$UpdateTestGroupConfigFileName"
	$testGroupConfigContent = [Xml] (Get-Content $testGroupFilePath)
	$addNode = $testGroupConfigContent.Configuration.Add
	$addNode.OfficeClientEdition = $bitness
	$addNode.Version = $version

	$updatesNode = $testGroupConfigContent.Updates
	$updatesNode.UpdatePath = $path
	$updatesNode.TaregtVersion = $version

	$testGroupConfigContent.Save($testGroupFilePath)

	$setupExePath = "$path\$c2rFileName"

	#2 - Run Setup.exe to download bits for specified version

	#Connect PowerShell to Share location
	$startLocation = Get-Location
	Set-Location $path
	# set up the executable with appropriate arguments
	$app = ".\$c2rFileName" 
	$arguments = "/download", "$UpdateSourceConfigFileName"

	#run the executable, this will trigger the download of bits to \\ShareName\Office\Data\
	& $app @arguments

	Set-Location $startLocation

	#connect to sccm PowerShell
	Set-Location "$siteId`:"

	# The SCCM Script runs here
}








