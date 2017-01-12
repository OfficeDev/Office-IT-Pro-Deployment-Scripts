. .\Generate-ODTConfigurationXML.ps1

$tempPath = "$env:TEMP\localConfig.xml"

Generate-ODTConfigurationXML -TargetFilePath $tempPath
