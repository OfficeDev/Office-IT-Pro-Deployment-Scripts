[CmdletBinding(SupportsShouldProcess=$true)]
Param
(
	[Parameter(Mandatory=$True)]
	[String]$UncPath,
	
	[Parameter()]
	[String]$Bitness = '32'
)
Begin
{
	$currentExecutionPolicy = Get-ExecutionPolicy
	Set-ExecutionPolicy Unrestricted -Scope Process -Force  
    $startLocation = Get-Location
}
Process
{
	Write-Host 'Updating Config Files'
	
	$setupFileName = 'setup.exe'
	$localSetupFilePath = ".\$setupFileName"
	$setupFilePath = "$UncPath\$localSetupFilePath"
	
	Copy-Item -Path $localSetupFilePath -Destination $UncPath -Force
	
	$downloadConfigFileName = 'Configure_Download.xml'
	$downloadConfigFilePath = "$UncPath\$downloadConfigFileName"
	$localDownloadConfigFilePath = ".\$downloadConfigFileName"
	
	$installConfigFileName = 'Configuration_InstallLocally.xml'
	$installConfigFilePath = "$UncPath\$installConfigFileName"
	$localInstallConfigFilePath = ".\$installConfigFileName"
	
	$content = [Xml](Get-Content $localDownloadConfigFilePath)
	$addNode = $content.Configuration.Add
	$addNode.OfficeClientEdition = $Bitness
	Write-Host 'Saving Download Configuration XML'	
	$content.Save($downloadConfigFilePath)
	
	$content = [Xml](Get-Content $localInstallConfigFilePath)
	$addNode = $content.Configuration.Add
	$addNode.OfficeClientEdition = $Bitness
	$addNode.SourcePath = $UncPath
	$updatesNode = $content.Configuration.Updates
	$updatesNode.UpdatePath = $UncPath
	Write-Host 'Saving Install Configuration XML'
	$content.Save($installConfigFilePath)
	
	Write-Host 'Setting up Click2Run to download Office to UNC Path'
	
	Set-Location $UncPath
	
	$app = ".\$setupFileName"
	$arguments = "/download", "$downloadConfigFileName"
	
	Write-Host 'Starting Download'
	& $app @arguments
	
	Write-Host 'Download Complete'	
}
End
{
	Set-ExecutionPolicy $currentExecutionPolicy -Scope Process -Force
    Set-Location $startLocation
}