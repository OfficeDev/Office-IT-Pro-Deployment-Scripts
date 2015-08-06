
Param
(
    [Parameter(Mandatory=$True)]
    [string]$UncPath,
    [Parameter(Mandatory=$True)]
    [string]$CommonFileShare
)

# Return the bitness of Windows.
function Test-64BitOS {
    [Management.ManagementBaseObject] $os = Get-WMIObject win32_operatingsystem
    if ($os.OSArchitecture -match "^64")
    {
        return $true
    }
    return $false
}

# Returns the version of Office
function Get-OfficeVersion {
    $objExcel = New-Object -ComObject Excel.Application
    return $objExcel.Version
}

# Add a registry key given its name, value and type.
function Add-RegistryKey( `
    [string] $key, `
    [string] $name, `
    [string] $value, `
    [string] $type)
{
    if (-not (Test-Path $key))
    {
        New-Item $key -Force | Out-Null
    }
    New-ItemProperty $key -Name $name -Value $value -PropertyType $type -Force | Out-Null
}

function Run-AgentInstaller {
    Copy-Item -Path "$UncPath\*" -Destination $env:temp

    if (!(Test-64BitOS))
    {
    Start-Process -FilePath "$env:Temp\osmia32.msi"
    }
    else
    {
    Start-Process -FilePath "$env:Temp\osmia64.msi"
    }
}


# Set the registry values to enable Telemetry Agent to upload data.
function Configure-TelemetryAgent([string] $database, [string] $folderName) {

    $objExcel = New-Object -ComObject Excel.Application
    $officeVersion = $objExcel.Version.Split('.')[0]

    if($officeVersion -match '11'){
        $key = "HKCU:\Software\Policies\Microsoft\Office\11.0\osm"
    }
    elseif($officeVersion -match '12'){
        $key = "HKCU:\Software\Policies\Microsoft\Office\12.0\osm"
    }
    elseif($officeVersion -match '14'){
        $key = "HKCU:\Software\Policies\Microsoft\Office\14.0\osm"
    }
    elseif($officeVersion -match '15'){
        $key = "HKCU:\Software\Policies\Microsoft\Office\15.0\osm"
    }
    elseif($officeVersion -match '16'){
        $key = "HKCU:\Software\Policies\Microsoft\Office\16.0\osm"
    }
    
    Add-RegistryKey $key "CommonFileShare" "$CommonFileShare"  "String"

    Add-RegistryKey $key "Tag1" "TAG1" "String"
    Add-RegistryKey $key "Tag2" "TAG2" "String"
    Add-RegistryKey $key "Tag3" "TAG3" "String"
    Add-RegistryKey $key "Tag4" "TAG4" "String"

    Add-RegistryKey $key "AgentInitWait" "1" "DWord"
    Add-RegistryKey $key "Enablelogging" "1" "DWord"
    Add-RegistryKey $key "EnableUpload" "1" "DWord"
    Add-RegistryKey $key "EnableFileObfuscation" "0" "DWord"
    Add-RegistryKey $key "AgentRandomDelay" "0" "DWord"
    
}


Run-AgentInstaller

Configure-TelemetryAgent