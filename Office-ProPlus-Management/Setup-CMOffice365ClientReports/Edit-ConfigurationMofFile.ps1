function Edit-ConfigurationMofFile{
<#
.Synopsis
Edits the configuration.mof file.

.DESCRIPTION
Appends the configuration.mof file in the Configuration Manager installation directory with the content of configuration.txt.

.PARAMETER configMofTxt
The name of the configuration.txt file.

.EXAMPLE
Edit-ConfigurationMofFile
Appends the content inside of configuration.txt into the configuration.mof file.
#>
Param(
    [Parameter(Mandatory=$false)]
    $configMofTxt = "configuration.txt",

    [System.Management.Automation.PSCredential]$Credentials
)

Begin{
    $ConfigMofFileName = "configuration.mof"
    $configMofLocation = "inboxes\clifiles.src\hinv"
    
    
    $HKLM = [UInt32] "0x80000002"
    $installKey = 'SOFTWARE\Microsoft\SMS'

    if($Credentials) {
       $regProv = Get-Wmiobject -list "StdRegProv" -namespace root\default -computername $env:COMPUTERNAME -Credential $Credentials
    } 
    else{
       $regProv = Get-Wmiobject -list "StdRegProv" -namespace root\default -computername $env:COMPUTERNAME
    }
}

Process{

    #Check uninstall path for ConfigMgr version   
    $ConfigMgrVersion = $regProv.EnumKey($HKLM, $installKey)
    foreach ($key in $ConfigMgrVersion.sNames) {
        if($key -match 'Setup'){
            $path = Join-Path $installKey $key
            $configItems = $regProv.EnumValues($HKLM, $path)
            foreach($item in $configItems.sNames){
                if($item -eq 'Installation Directory'){
                    $installationDirectory = $regProv.GetStringValue($HKLM, $path, $item).sValue
                }
            }
        }
    }

    $ConfigurationMofPath = Join-Path $installationDirectory $configMofLocation
    $configMofFilePath = $ConfigurationMofPath + "\" + $ConfigMofFileName
    $backupMofFilePath = $configMofFilePath + ".backup"
  
    #Make a backup of the mof file
    Copy-Item -Path $configMofFilePath -Destination $backupMofFilePath

    #Append the mof file
    Add-Content -Path $configMofFilePath -Value (Get-Content $configMofTxt)
}
}
