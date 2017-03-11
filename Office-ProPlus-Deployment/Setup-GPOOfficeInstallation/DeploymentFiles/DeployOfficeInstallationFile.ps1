param(
    [Parameter()]
    [string]$OfficeDeploymentPath,
    
	[Parameter(Mandatory=$true)]
	[String]$OfficeDeploymentFileName = $NULL,

    [Parameter()]
    [string]$Quiet = "True"
)

Set-Location $OfficeDeploymentPath

$scriptPath = "."
. $scriptPath\SharedFunctions.ps1

$officeProducts = Get-OfficeVersion -ShowAllInstalledProducts | Select *
$Office2016C2RExists = $officeProducts | Where {$_.ClickToRun -eq $true -and $_.Version -like '16.*' }

if(!$Office2016C2RExists){
    $ActionFile = "$OfficeDeploymentPath\$OfficeDeploymentFileName"

    if($OfficeDeploymentFileName.EndsWith("msi")){
        if($Quiet -eq "True"){
            $argList = "/qn /norestart"
        } else {
            $argList = "/norestart"
        }

        $cmdLine = """$ActionFile"" $argList"
        $cmd = "cmd /c msiexec /i $cmdLine"
    } elseif($OfficeDeploymentFileName.EndsWith("exe")){
        if($Quiet -eq "True"){
            $argList = "/silent"
        }

        $cmd = "$ActionFile $argList"
    }

    Invoke-Expression $cmd
}