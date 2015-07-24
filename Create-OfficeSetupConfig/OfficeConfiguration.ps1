
Function New-OfficeConfiguration{
    [OutputType([string])]
    Param(

    [Parameter(Mandatory=$true)]
    [string] $Bitness,

    [Parameter(Mandatory=$true)]
    [string] $ProductId,

    [Parameter()]
    [string] $LanguageId = (Get-Culture | %{$_.Name}),

    [Parameter(Mandatory=$true)]
    [string] $OutPath

    )

    Process{
        #Create Document and Add root Configuration Element
        [System.XML.XMLDocument]$configFile = New-Object System.XML.XMLDocument
        [System.XML.XMLElement]$ConfigurationRoot=$configFile.CreateElement("Configuration")
        $configFile.appendChild($ConfigurationRoot) | Out-Null

        #Add the Add Element under Configuration and set the Bitness
        [System.XML.XMLElement]$AddElement=$configFile.CreateElement("Add")
        $ConfigurationRoot.appendChild($AddElement) | Out-Null
        $AddElement.SetAttribute("OfficeClientEdition",$Bitness) | Out-Null

        #Add the Product Element under Add and set the ID
        [System.XML.XMLElement]$ProductElement=$configFile.CreateElement("Product")
        $AddElement.appendChild($ProductElement) | Out-Null
        $ProductElement.SetAttribute("ID",$ProductId) | Out-Null

        #Add the Language Element under Product and set the ID
        [System.XML.XMLElement]$LanguageElement=$configFile.CreateElement("Language")
        $ProductElement.appendChild($LanguageElement) | Out-Null
        $LanguageElement.SetAttribute("ID",$LanguageId) | Out-Null
        $configFile.Save($OutPath) | Out-Null

        return $OutPath;
    }
}

Function Remove-Product{

    Param(

        [Parameter()]
        [switch] $All,

        [Parameter()]
        [string] $ProductId,

        [Parameter()]
        [string[]] $LanguageIds,

        [Parameter(Mandatory=$true)]
        [string] $ConfigPath

    )

    Process{

        [System.XML.XMLDocument]$configFile = New-Object System.XML.XMLDocument
        $configFile.Load($ConfigPath) | Out-Null

        if($configFile.Configuration -eq $null){
            throw $NoConfigurationElement
        }

        if($configFile.COnfiguration.Remove -eq $null){
            [System.XML.XMLElement]$RemoveElement=$configFile.CreateElement("Remove")
            $configFile.COnfiguration.appendChild($RemoveElement) | Out-Null
        }
        if($All){
            $configFile.COnfiguration.Remove.SetAttribute("All", "True") | Out-Null
        }else{
            [System.XML.XMLElement]$ProductElement = $configFile.Configuration.Remove.Product | ?  ID -eq $ProductId
            if($ProductElement -eq $null){
                [System.XML.XMLElement]$ProductElement=$configFile.CreateElement("Product")
                $configFile.COnfiguration.Remove.appendChild($ProductElement) | Out-Null
                $ProductElement.SetAttribute("ID", $ProductId) | Out-Null
            }
            foreach($LanguageId in $LanguageIds){
                [System.XML.XMLElement]$LanguageElement = $configFile.Configuration.Remove.Product.Language | ?  ID -eq $LanguageId
                if($LanguageElement -eq $null){
                    [System.XML.XMLElement]$LanguageElement=$configFile.CreateElement("Language")
                    $ProductElement.appendChild($LanguageElement) | Out-Null
                    $LanguageElement.SetAttribute("ID", $LanguageId) | Out-Null
                }
            }
        }

        $configFile.Save($ConfigPath) | Out-Null

        return $ConfigPath
    }

}

Function Add-Product{

    Param(

        [Parameter()]
        [string] $ProductId,

        [Parameter()]
        [string[]] $LanguageIds,

        [Parameter()]
        [string[]] $ExcludeApps,

        [Parameter(Mandatory=$true)]
        [string] $ConfigPath

    )

    Process{

        [System.XML.XMLDocument]$configFile = New-Object System.XML.XMLDocument
        $configFile.Load($ConfigPath) | Out-Null

        if($configFile.Configuration -eq $null){
            throw $NoConfigurationElement
        }

        if($configFile.COnfiguration.Add -eq $null){
            throw $NoAddElement
        }

        [System.XML.XMLElement]$ProductElement = $configFile.Configuration.Add.Product | ?  ID -eq $ProductId
        if($ProductElement -eq $null){
            [System.XML.XMLElement]$ProductElement=$configFile.CreateElement("Product")
            $configFile.COnfiguration.Remove.appendChild($ProductElement) | Out-Null
            $ProductElement.SetAttribute("Id", $ProductId) | Out-Null
        }


        foreach($LanguageId in $LanguageIds){
            [System.XML.XMLElement]$LanguageElement = $configFile.Configuration.Add.Product.Language | ?  ID -eq $LanguageId
            if($LanguageElement -eq $null){
                [System.XML.XMLElement]$LanguageElement=$configFile.CreateElement("Language")
                $ProductElement.appendChild($LanguageElement) | Out-Null
                $LanguageElement.SetAttribute("ID", $LanguageId) | Out-Null
            }
        }

        foreach($ExcludeApp in $ExcludeApps){
            [System.XML.XMLElement]$ExcludeAppElement = $configFile.Configuration.Add.Product.ExcludeApp | ?  ID -eq $ExcludeApp
            if($ExcludeAppElement -eq $null){
                [System.XML.XMLElement]$ExcludeAppElement=$configFile.CreateElement("ExcludeApp")
                $ProductElement.appendChild($ExcludeAppElement) | Out-Null
                $ExcludeAppElement.SetAttribute("ID", $ExcludeApp) | Out-Null
            }
        }

        $configFile.Save($ConfigPath) | Out-Null

        return $ConfigPath
    }

}

Function Set-Updates{

    Param(

        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [string] $ConfigPath,
        
        [Parameter()]
        [string] $Enabled,

        [Parameter()]
        [string] $UpdatePath,

        [Parameter()]
        [string] $TargetVersion,

        [Parameter()]
        [string] $Deadline

    )

    Process{

        [System.XML.XMLDocument]$configFile = New-Object System.XML.XMLDocument
        $configFile.Load($ConfigPath) | Out-Null

        if($configFile.Configuration -eq $null){
            throw $NoConfigurationElement
        }

        if($configFile.COnfiguration.Updates -eq $null){
            [System.XML.XMLElement]$UpdateElement=$configFile.CreateElement("Updates")
            $configFile.COnfiguration.appendChild($UpdateElement) | Out-Null
        }

        if([string]::IsNullOrWhiteSpace($Enabled) -eq $false){
            $configFile.COnfiguration.Updates.SetAttribute("Enabled", $Enabled) | Out-Null
        }
        if([string]::IsNullOrWhiteSpace($UpdatePath) -eq $false){
            $configFile.COnfiguration.Updates.SetAttribute("UpdatePath", $UpdatePath) | Out-Null
        }
        if([string]::IsNullOrWhiteSpace($TargetVersion) -eq $false){
            $configFile.COnfiguration.Updates.SetAttribute("TargetVersion", $TargetVersion) | Out-Null
        }
        if([string]::IsNullOrWhiteSpace($Deadline) -eq $false){
            $configFile.COnfiguration.Updates.SetAttribute("Deadline", $Deadline) | Out-Null
        }

        $configFile.Save($ConfigPath) | Out-Null

        return $ConfigPath
    }
}

Function Set-ConfigProperties{

    Param(
        
        [Parameter()]
        [string] $AutoActivate,

        [Parameter()]
        [string] $ForceAppShutDown,

        [Parameter()]
        [string] $PackageGUID,

        [Parameter()]
        [string] $SharedComputerLicensing,

        [Parameter()]
        [string] $ConfigPath
    )

    Process{
        [System.XML.XMLDocument]$configFile = New-Object System.XML.XMLDocument
        $configFile.Load($ConfigPath) | Out-Null

        if($configFile.Configuration -eq $null){
            throw $NoConfigurationElement
        }

        if([string]::IsNullOrWhiteSpace($AutoActivate) -eq $false){
            [System.XML.XMLElement]$AutoActivateElement = $configFile.Configuration.Property | ?  Name -eq "AUTOACTIVATE"
            if($AutoActivateElement -eq $null){
                [System.XML.XMLElement]$AutoActivateElement=$configFile.CreateElement("Property")
            }
                
            $configFile.COnfiguration.appendChild($AutoActivateElement) | Out-Null
            $AutoActivateElement.SetAttribute("Name", "AUTOACTIVATE") | Out-Null
            $AutoActivateElement.SetAttribute("Value", $AutoActivate) | Out-Null
        }

        if([string]::IsNullOrWhiteSpace($ForceAppShutDown) -eq $false){
            [System.XML.XMLElement]$ForceAppShutDownElement = $configFile.Configuration.Property | ?  Name -eq "FORCEAPPSHUTDOWN"
            if($ForceAppShutDownElement -eq $null){
                [System.XML.XMLElement]$ForceAppShutDownElement=$configFile.CreateElement("Property")
            }
                
            $configFile.COnfiguration.appendChild($ForceAppShutDownElement) | Out-Null
            $ForceAppShutDownElement.SetAttribute("Name", "FORCEAPPSHUTDOWN") | Out-Null
            $ForceAppShutDownElement.SetAttribute("Value", $ForceAppShutDownElement) | Out-Null
        }

        if([string]::IsNullOrWhiteSpace($PackageGUID) -eq $false){
            [System.XML.XMLElement]$PackageGUIDElement = $configFile.Configuration.Property | ?  Name -eq "PACKAGEGUID"
            if($PackageGUIDElement -eq $null){
                [System.XML.XMLElement]$PackageGUIDElement=$configFile.CreateElement("Property")
            }
                
            $configFile.COnfiguration.appendChild($PackageGUIDElement) | Out-Null
            $PackageGUIDElement.SetAttribute("Name", "PACKAGEGUID") | Out-Null
            $PackageGUIDElement.SetAttribute("Value", $PackageGUID) | Out-Null
        }

        if([string]::IsNullOrWhiteSpace($SharedComputerLicensing) -eq $false){
            [System.XML.XMLElement]$SharedComputerLicensingElement = $configFile.Configuration.Property | ?  Name -eq "SharedComputerLicensing"
            if($SharedComputerLicensingElement -eq $null){
                [System.XML.XMLElement]$SharedComputerLicensingElement=$configFile.CreateElement("Property")
            }
                
            $configFile.COnfiguration.appendChild($SharedComputerLicensingElement) | Out-Null
            $SharedComputerLicensingElement.SetAttribute("Name", "SharedComputerLicensing") | Out-Null
            $SharedComputerLicensingElement.SetAttribute("Value", $SharedComputerLicensing) | Out-Null
        }

        $configFile.Save($ConfigPath) | Out-Null
        return $ConfigPath
        
    }
}

Function Set-Add{

    Param(

        [Parameter()]
        [string] $SourcePath,

        [Parameter()]
        [string] $Version,

        [Parameter()]
        [string] $Bitness,

        [Parameter()]
        [string] $ConfigPath

    )

    Process{

        [System.XML.XMLDocument]$configFile = New-Object System.XML.XMLDocument
        $configFile.Load($ConfigPath) | Out-Null

        if($configFile.Configuration -eq $null){
            throw $NoConfigurationElement
        }

        if($configFile.COnfiguration.Add -eq $null){
            [System.XML.XMLElement]$AddElement=$configFile.CreateElement("Add")
            $configFile.COnfiguration.appendChild($AddElement) | Out-Null
        }

        if([string]::IsNullOrWhiteSpace($SourcePath) -eq $false){
            $configFile.COnfiguration.Add.SetAttribute("SourcePath", $SourcePath) | Out-Null
        }
        if([string]::IsNullOrWhiteSpace($Version) -eq $false){
            $configFile.COnfiguration.Add.SetAttribute("Version", $Version) | Out-Null
        }
        if([string]::IsNullOrWhiteSpace($Bitness) -eq $false){
            $configFile.COnfiguration.Add.SetAttribute("OfficeClientEdition", $Bitness) | Out-Null
        }

        $configFile.Save($ConfigPath) | Out-Null

        return $ConfigPath

    }

}

Function Set-Logging{
    
    Param(

        [Parameter()]
        [string] $Level,

        [Parameter()]
        [string] $Path,

        [Paramter()]
        [string] $ConfigPath

    )

    Process{

        [System.XML.XMLDocument]$configFile = New-Object System.XML.XMLDocument
        $configFile.Load($ConfigPath) | Out-Null

        if($configFile.Configuration -eq $null){
            throw $NoConfigurationElement
        }

        if($configFile.COnfiguration.Logging -eq $null){
            [System.XML.XMLElement]$LoggingElement=$configFile.CreateElement("Logging")
            $configFile.COnfiguration.appendChild($LoggingElement) | Out-Null
        }

        if([string]::IsNullOrWhiteSpace($Level) -eq $false){
            $configFile.COnfiguration.Logging.SetAttribute("Level", $Level) | Out-Null
        }
        if([string]::IsNullOrWhiteSpace($Path) -eq $false){
            $configFile.COnfiguration.Logging.SetAttribute("Path", $Path) | Out-Null
        }

        $configFile.Save($ConfigPath) | Out-Null

        return $ConfigPath

    }
}

Function Set-Display{

    Param(

        [Parameter()]
        [string] $Level,

        [Parameter()]
        [string] $AcceptEULA,

        [Parameter()]
        [string] $ConfigPath

    )

    Process{

        [System.XML.XMLDocument]$configFile = New-Object System.XML.XMLDocument
        $configFile.Load($ConfigPath) | Out-Null

        if($configFile.Configuration -eq $null){
            throw $NoConfigurationElement
        }

        if($configFile.COnfiguration.Display -eq $null){
            [System.XML.XMLElement]$DisplayElement=$configFile.CreateElement("Display")
            $configFile.COnfiguration.appendChild($DisplayElement) | Out-Null
        }

        if([string]::IsNullOrWhiteSpace($Level) -eq $false){
            $configFile.COnfiguration.Display.SetAttribute("Level", $Level) | Out-Null
        }
        if([string]::IsNullOrWhiteSpace($Path) -eq $AcceptEULA){
            $configFile.COnfiguration.Display.SetAttribute("AcceptEULA", $AcceptEULA) | Out-Null
        }

        $configFile.Save($ConfigPath) | Out-Null

        return $ConfigPath

    }

}