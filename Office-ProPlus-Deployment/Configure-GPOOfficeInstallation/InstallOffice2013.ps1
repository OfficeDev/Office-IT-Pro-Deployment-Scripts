[CmdletBinding(SupportsShouldProcess=$true)]
Param
(
	[Parameter(Mandatory=$True)]
	[String]$UncPath,
	
	[Parameter(Mandatory=$True)]
	[String]$ConfigFileName
)

Set-Location $UncPath

$c2RFileName = "SetupOffice2013.exe"
$app = ".\$c2RFileName"

$arguments = "/configure", "$ConfigFileName"

& $app @arguments