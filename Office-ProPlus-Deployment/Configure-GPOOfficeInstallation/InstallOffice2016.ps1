[CmdletBinding(SupportsShouldProcess=$true)]
Param
(
	[Parameter(Mandatory=$True)]
	[String]$UncPath,
	
	[Parameter(Mandatory=$True)]
	[String]$ConfigFileName
)

Set-Location $UncPath

$c2RFileName = "Office2016Setup.exe"
$app = ".\$c2RFileName"

$arguments = "/configure", "$ConfigFileName"

& $app @arguments