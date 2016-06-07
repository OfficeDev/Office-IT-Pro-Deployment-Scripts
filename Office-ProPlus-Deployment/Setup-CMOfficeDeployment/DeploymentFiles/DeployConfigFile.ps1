  param(
	[Parameter(Mandatory=$true)]
	[String]$ConfigFileName = $NULL
  )

Process {
 $scriptPath = "."

 if ($PSScriptRoot) {
   $scriptPath = $PSScriptRoot
 } else {
   $scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
 }

 $shareFunctionsPath = "$scriptPath\SharedFunctions.ps1"
 if ($scriptPath.StartsWith("\\")) {
 } else {
    if (!(Test-Path -Path $shareFunctionsPath)) {
        throw "Missing Dependency File SharedFunctions.ps1"    
    }
 }
 . $shareFunctionsPath

  #Importing all required functions
. $scriptPath\Edit-OfficeConfigurationFile.ps1
. $scriptPath\Install-OfficeClickToRun.ps1
. $scriptPath\SharedFunctions.ps1

 [string]$configFilePath = "$scriptPath\$ConfigFileName"
 [string]$targetFilePath = "$env:temp\configuration.xml"

 if (!(Test-Path -Path $configFilePath)) {
     throw "Cannot find Configuration Xml File: $ConfigFileName"
 }

 Copy-Item -Path $configFilePath -Destination $targetFilePath -Force

 [string]$UpdateSource = (Get-ODTAdd -TargetFilePath $targetFilePath | select SourcePath).SourcePath
 [string]$Bitness = (Get-ODTAdd -TargetFilePath $targetFilePath | select OfficeClientEdition).OfficeClientEdition
 if($Bitness -eq '64'){
    $Bitness = "x64"
 }
 else{
    $Bitness = "x32"
 }

 if ($UpdateSource) {
     if ($UpdateSource.StartsWith(".\")) {
        $UpdateSource = $UpdateSource -replace "^\.", "$scriptPath"
     }
 }

 $UpdateURLPath = $NULL
 if ($UpdateSource) {
   if (Test-ItemPathUNC -Path "$UpdateSource") {
      $UpdateURLPath = "$UpdateURLPath\$SourceFileFolder"
   }
 }

$languages = Get-XMLLanguages -Path $targetFilePath

if ($UpdateSource) {
    $ValidUpdateSource = Test-UpdateSource -UpdateSource $UpdateSource -OfficeLanguages $languages -Bitness $Bitness
    if ($ValidUpdateSource) {
       Set-ODTAdd -TargetFilePath $targetFilePath -SourcePath $UpdateSource | Out-Null
    } else {
       throw "Invalid Update Source: $UpdateSource"
    }
}

Install-OfficeClickToRun -TargetFilePath $targetFilePath

}