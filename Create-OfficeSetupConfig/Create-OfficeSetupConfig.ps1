<#
.SYNOPSIS
Short Description

.DESCRIPTION
Long Description

.PARAMETER ProductId
Explaination of ProductId

.PARAMETER AddLanguageOptions
Explaination of AddLanguageOptions

.PARAMETER ARPOptions
Explaination of ARPOptions

.PARAMETER CommandOptions
Explaination of CommandOptions

.PARAMETER CompanyName
Explaination of CompanyName

.PARAMETER DisplayOptions
Explaination of DisplayOptions

.PARAMETER DistributionPointPath
Explaination of DistributionPointPath

.PARAMETER InstallLocation
Explaination of InstallLocation

.PARAMETER LISOptions
Explaination of LISOptions

.PARAMETER LoggingOptions
Explaination of LoggingOptions

.PARAMETER OptionStateList
Explaination of OptionStateList

.PARAMETER PIDKEY
Explaination of PIDKEY

.PARAMETER RemoveLanguageOptions
Explaination of RemoveLanguageOptions

.PARAMETER SettingOptions
Explaination of SettingOptions

.PARAMETER SetupUpdatesOptions
Explaination of SetupUpdatesOptions

.PARAMETER UserInitials
Explaination of UserInitials

.PARAMETER Username
Explaination of Username

.Example
./Skeleton.ps1 -myParam1 "Value1" -myParam2 "Value2"
Usage example one

.Example
./Skeleton.ps1 -Param1 "Value1"
Usage example two

.Notes
Additional explaination. Long and indepth examples should also go here.

.Link
http://relevantlink.com

.Link
relevent-command

#>

[CmdletBinding()]
Param(

    [Parameter()]
    [string] $ProductId,

    [Parameter()]
    [Hashtable[]] $AddLanguageOptions,

    [Parameter()]
    [Hashtable] $ARPOptions,

    [Parameter()]
    [Hashtable] $CommandOptions,

    [Parameter()]
    [string] $CompanyName,

    [Parameter()]
    [Hashtable] $DisplayOptions,

    [Parameter()]
    [string] $DistributionPointPath,

    [Parameter()]
    [string] $InstallLocation,

    [Parameter()]
    [Hashtable] $LISOptions,

    [Parameter()]
    [Hashtable] $LoggingOptions,

    [Parameter()]
    [Hashtable[]] $OptionStateList,

    [Parameter()]
    [string] $PIDKEY,

    [Parameter()]
    [string[]] $RemoveLanguageOptions,

    [Parameter()]
    [Hashtable[]] $SettingOptions,

    [Parameter()]
    [Hashtable] $SetupUpdatesOptions,

    [Parameter()]
    [string] $UserInitials,

    [Parameter()]
    [string] $Username

)

#main script flow
[System.XML.XMLDocument]$configFile = New-Object System.XML.XMLDocument
[System.XML.XMLElement]$ConfigurationRoot=$configFile.CreateElement("Configuration")
$configFile.appendChild($ConfigurationRoot)
$ConfigurationRoot.SetAttribute("Product",$ProductId)


if($CommandOptions -ne $Null){
    foreach($CommandOption in $CommandOptions){
        [System.XML.XMLElement]$CommandElement=$configFile.CreateElement("Command")
        $configFile.Configuration.appendChild($CommandElement)

        $CommandElement.SetAttribute("Path",$CommandOption.Path)

        if([String]::IsNullOrWhiteSpace($CommandOption.QuietArgs) -eq $false){
            $CommandElement.SetAttribute("QuietArgs",$CommandOption.QuietArgs)
        }

        if([String]::IsNullOrWhiteSpace($CommandOption.Args) -eq $false){
            $CommandElement.SetAttribute("Args",$CommandOption.Args)
        }

        if([String]::IsNullOrWhiteSpace($CommandOption.ChainPosition) -eq $false){
            $CommandElement.SetAttribute("ChainPosition",$CommandOption.ChainPosition)
        }else{
            $CommandElement.SetAttribute("ChainPosition","After")
        }

        if([String]::IsNullOrWhiteSpace($CommandOption.Wait) -eq $false){
            $CommandElement.SetAttribute("Wait",$CommandOption.Wait)
        }

        if([String]::IsNullOrWhiteSpace($CommandOption.Execute) -eq $false){
            $CommandElement.SetAttribute("Execute",$CommandOption.Execute)
        }else{
            $CommandElement.SetAttribute("Execute","Install")
        }

        if([String]::IsNullOrWhiteSpace($CommandOption.Platform) -eq $false){
            $CommandElement.SetAttribute("Platform",$CommandOption.Platform)
        }else{
            $CommandElement.SetAttribute("Platform", "x86")
        }
    }
}

if($AddLanguageOptions -ne $Null){
    foreach($AddLanguageOption in $AddLanguageOptions){
        [System.XML.XMLElement]$AddLanguageElement=$configFile.CreateElement("AddLanguage")
        $configFile.Configuration.appendChild($AddLanguageElement)

        if([String]::IsNullOrWhiteSpace($AddLanguageOption.Id) -eq $false){
            $AddLanguageElement.SetAttribute("Id",$AddLanguageOption.Id)
        }

        if([String]::IsNullOrWhiteSpace($AddLanguageOption.ShellTransform) -eq $false){
            $AddLanguageElement.SetAttribute("ShellTransform",$AddLanguageOption.ShellTransform)
        }
    }
}

if($ARPOptions -ne $Null){
    [System.XML.XMLElement]$ARPElement=$configFile.CreateElement("ARP")
    $configFile.Configuration.appendChild($ARPElement)

    if([String]::IsNullOrWhiteSpace($ARPOptions.ARPCOMMENTS) -eq $false){
        $ARPElement.SetAttribute("ARPCOMMENTS",$ARPOptions.ARPCOMMENTS)
    }

    if([String]::IsNullOrWhiteSpace($ARPOptions.ARPCONTACT) -eq $false){
        $ARPElement.SetAttribute("ARPCONTACT",$ARPOptions.ARPCONTACT)
    }

    if([String]::IsNullOrWhiteSpace($ARPOptions.ARPNOMODIFY) -eq $false){
        $ARPElement.SetAttribute("ARPNOMODIFY",$ARPOptions.ARPNOMODIFY)
    }

    if([String]::IsNullOrWhiteSpace($ARPOptions.ARPNOREMOVE) -eq $false){
        $ARPElement.SetAttribute("ARPNOREMOVE",$ARPOptions.ARPNOREMOVE)
    }

    if([String]::IsNullOrWhiteSpace($ARPOptions.ARPURLINFOABOUT) -eq $false){
        $ARPElement.SetAttribute("ARPURLINFOABOUT",$ARPOptions.ARPURLINFOABOUT)
    }

    if([String]::IsNullOrWhiteSpace($ARPOptions.ARPURLUPDATEINFO) -eq $false){
        $ARPElement.SetAttribute("ARPURLUPDATEINFO",$ARPOptions.ARPURLUPDATEINFO)
    }

    if([String]::IsNullOrWhiteSpace($ARPOptions.ARPHELPLINK) -eq $false){
        $ARPElement.SetAttribute("ARPHELPLINK",$ARPOptions.ARPHELPLINK)
    }

    if([String]::IsNullOrWhiteSpace($ARPOptions.ARPHELPTELEPHONE) -eq $false){
        $ARPElement.SetAttribute("ARPHELPTELEPHONE",$ARPOptions.ARPHELPTELEPHONE)
    }
}

if([string]::IsNullOrWhiteSpace($CompanyName) -eq $false){
    [System.XML.XMLElement]$CompanyNameElement=$configFile.CreateElement("COMPANYNAME")
    $configFile.Configuration.appendChild($CompanyNameElement)
    $CompanyNameElement.SetAttribute("Value",$CompanyName)
}

if($DisplayOptions -ne $Null){
    [System.XML.XMLElement]$DisplayElement=$configFile.CreateElement("Display")
    $configFile.Configuration.appendChild($DisplayElement)

    if([String]::IsNullOrWhiteSpace($DisplayOptions.Level) -eq $false){
        $DisplayElement.SetAttribute("Level",$DisplayOptions.Level)
    }else{
        $DisplayElement.SetAttribute("Level","Full")
    }

    if([String]::IsNullOrWhiteSpace($DisplayOptions.CompletionNotice) -eq $false){
        $DisplayElement.SetAttribute("CompletionNotice",$DisplayOptions.CompletionNotice)
    }else{
        $DisplayElement.SetAttribute("CompletionNotice","No")
    }

    if([String]::IsNullOrWhiteSpace($DisplayOptions.SuppressModal) -eq $false){
        $DisplayElement.SetAttribute("SuppressModal",$DisplayOptions.SuppressModal)
    }else{
        $DisplayElement.SetAttribute("SuppressModal","No")
    }

    if([String]::IsNullOrWhiteSpace($DisplayOptions.NoCancel) -eq $false){
        $DisplayElement.SetAttribute("NoCancel",$DisplayOptions.NoCancel)
    }else{
        $DisplayElement.SetAttribute("NoCancel","No")
    }

    if([String]::IsNullOrWhiteSpace($DisplayOptions.AcceptEula) -eq $false){
        $DisplayElement.SetAttribute("AcceptEula",$DisplayOptions.AcceptEula)
    }else{
        $DisplayElement.SetAttribute("AcceptEula","No")
    }
}

if([string]::IsNullOrWhiteSpace($DistributionPointPath) -eq $false){
        [System.XML.XMLElement]$DistributionPointElement=$configFile.CreateElement("DistributionPoint")
        $configFile.Configuration.appendChild($DistributionPointElement)
        $DistributionPointElement.SetAttribute("Location",$DistributionPointPath)
}

if([string]::IsNullOrWhiteSpace($InstallLocation) -eq $false){
        [System.XML.XMLElement]$InstallLocationElement=$configFile.CreateElement("INSTALLLOCATION")
        $configFile.Configuration.appendChild($InstallLocationElement)
        $InstallLocationElement.SetAttribute("Value",$InstallLocation)
}

if($LISOptions -ne $Null){
        [System.XML.XMLElement]$LISElement=$configFile.CreateElement("LIS")
        $configFile.Configuration.appendChild($LISElement)

        if([String]::IsNullOrWhiteSpace($options.CACHEACTION) -eq $false){
            $LISElement.SetAttribute("CACHEACTION",$LISOptions.CACHEACTION)
        }

        if([String]::IsNullOrWhiteSpace($options.SOURCELIST) -eq $false){
            $LISElement.SetAttribute("SOURCELIST",$LISOptions.SOURCELIST)
        }
}

if($LoggingOptions -ne $Null){
    [System.XML.XMLElement]$LoggingElement=$configFile.CreateElement("Logging")
    $configFile.Configuration.appendChild($LoggingElement)

    if([String]::IsNullOrWhiteSpace($LoggingOptions.Type) -eq $false){
        $LoggingElement.SetAttribute("Type",$LoggingOptions.Type)
    }else{
        $LoggingElement.SetAttribute("Type", "Standard")
    }

    if([String]::IsNullOrWhiteSpace($LoggingOptions.Path) -eq $false){
        $LoggingElement.SetAttribute("Path",$LoggingOptions.Path)
    }

    if([String]::IsNullOrWhiteSpace($LoggingOptions.Template) -eq $false){
        $LoggingElement.SetAttribute("Template",$LoggingOptions.Template)
    }
}

if($OptionStateList -ne $Null){
    foreach($OptionState in $OptionStateList){
        [System.XML.XMLElement]$OptionElement=$configFile.CreateElement("OptionState")
        $configFile.Configuration.appendChild($LoggingElement)

        if([String]::IsNullOrWhiteSpace($OptionState.Id) -eq $false){
            $OptionElement.SetAttribute("Id",$OptionState.Id)
        }

        if([String]::IsNullOrWhiteSpace($OptionState.State) -eq $false){
            $OptionElement.SetAttribute("State",$OptionState.State)
        }

        if([String]::IsNullOrWhiteSpace($OptionState.Children) -eq $false){
            $OptionElement.SetAttribute("Template",$OptionState.Template)
        }
    }
}

if([string]::IsNullOrWhiteSpace($PIDKEY) -eq $false){
    [System.XML.XMLElement]$PIDKEYElement=$configFile.CreateElement("PIDKEY")
    $configFile.Configuration.appendChild($PIDKEYElement)
    $PIDKEYElement.SetAttribute("Value",$PIDKEY)
}

if($RemoveLanguageOptions -ne $Null){
    foreach($RemoveLanguageOption in $RemoveLanguageOptions){
        [System.XML.XMLElement]$RemoveLanguageElement=$configFile.CreateElement("RemoveLanguage")
        $configFile.Configuration.appendChild($RemoveLanguageElement)
        $RemoveLanguageElement.SetAttribute("Id",$RemoveLanguageOption)
    }
}

if($SettingOptions -ne $Null){
    foreach($Setting in $SettingOptions){
        [System.XML.XMLElement]$SettingElement=$configFile.CreateElement("Setting")
        $configFile.Configuration.appendChild($SettingElement)

        if([String]::IsNullOrWhiteSpace($Setting.Id) -eq $false){
            $SettingElement.SetAttribute("Id",$Setting.Id)
        }

        if([String]::IsNullOrWhiteSpace($Setting.Value) -eq $false){
            $SettingElement.SetAttribute("Value",$Setting.Value)
        }
    }
}

if($SetupUpdatesOptions -ne $Null){
    [System.XML.XMLElement]$SetupUpdatesElement=$configFile.CreateElement("SetupUpdates")
    $configFile.Configuration.appendChild($SetupUpdatesElement)

    if([String]::IsNullOrWhiteSpace($SetupUpdatesOptions.CheckForSUpdates) -eq $false){
        $SetupUpdatesElement.SetAttribute("CheckForSUpdates",$SetupUpdatesOptions.CheckForSUpdates)
    }

    if([String]::IsNullOrWhiteSpace($SetupUpdatesOptions.SUpdateLocation) -eq $false){
        $SetupUpdatesElement.SetAttribute("SUpdateLocation",$SetupUpdatesOptions.SUpdateLocation)
    }
}

if([string]::IsNullOrWhiteSpace($UserInitials) -eq $false){
    [System.XML.XMLElement]$UserInitialsElement=$configFile.CreateElement("USERINITIALS")
    $configFile.Configuration.appendChild($UserInitialsElement)
    $UserInitialsElement.SetAttribute("Value",$UserInitials)
}

if([string]::IsNullOrWhiteSpace($Username) -eq $false){
    [System.XML.XMLElement]$UserNameElement=$configFile.CreateElement("USERNAME")
    $configFile.Configuration.appendChild($UserNameElement)
    $UserNameElement.SetAttribute("Value",$Username)
}

$configFile.Save("$env:Public\Documents\config.xml")