    Param(
        [Parameter(Mandatory=$true)]
        [string] $GPOName,

        [Parameter()]
        [string] $VersionNumber = "15"
    )
    
    Process{
    #Set GPO Registry values given a GPOName for Telemetry
    $keyPath = "HKCU\Software\Policies\Microsoft\Office\$VersionNumber.0\osm"
    $GPO = Get-GPO -Name $GPOName

    Set-GPRegistryValue -Guid $GPO.Id -Key $keyPath -ValueName "CommonFileShare" -Type String -Value "$commonFileShare"

    Set-GPRegistryValue -Guid $GPO.Id -Key $keyPath -ValueName "Tag1" -Type String -Value "$commonFileShare"

    Set-GPRegistryValue -Guid $GPO.Id -Key $keyPath -ValueName "Tag2" -Type String -Value "$commonFileShare"

    Set-GPRegistryValue -Guid $GPO.Id -Key $keyPath -ValueName "Tag3" -Type String -Value "$commonFileShare"

    Set-GPRegistryValue -Guid $GPO.Id -Key $keyPath -ValueName "Tag4" -Type String -Value "$commonFileShare"

    Set-GPRegistryValue -Guid $GPO.Id -Key $keyPath -ValueName "AgentInitWait" -Type DWord -Value 00000001

    Set-GPRegistryValue -Guid $GPO.Id -Key $keyPath -ValueName "Enablelogging" -Type DWord -Value 00000001

    Set-GPRegistryValue -Guid $GPO.Id -Key $keyPath -ValueName "EnableUpload" -Type DWord -Value 00000001

    Set-GPRegistryValue -Guid $GPO.Id -Key $keyPath -ValueName "EnableFileObfuscation" -Type DWord -Value 00000000

    Set-GPRegistryValue -Guid $GPO.Id -Key $keyPath -ValueName "AgentRandomDelay" -Type DWord -Value 00000000

    }