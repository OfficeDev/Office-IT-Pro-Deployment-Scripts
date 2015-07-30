
$validProductIds = @("O365ProPlusRetail","O365BusinessRetail","VisioProRetail","ProjectProRetail", "SPDRetail")

Function New-ODTConfiguration{
<#
.SYNOPSIS
Creates a simple Office configuration file and outputs a 
string that is the path of the file

.DESCRIPTION
Given at least the bitness of the office version, the product id, and 
the file path of the output file, this function creates an xml file with
the bare minimum values to be usable. A configuration root, an add element,
a product element, and a language element (nested one after the other).
The output is the file path of the file so that this function can easily
be piped into the other associated functions. 

.PARAMETER Bitness
Possible values are '32' or '64'
Required. Specifies the edition of Click-to-Run for Office 365 product 
to use: 32- or 64-bit. The action fails if OfficeClientEdition is not 
set to a valid value.

A configure mode action may fail if OfficeClientEdition is set incorrectly. 
For example, if you attempt to install a 64-bit edition of a Click-to-Run 
for Office 365 product on a computer that is running a 32-bit Windows 
operating system, or if you try to install a 32-bit Click-to-Run for Office 
365 product on a computer that has a 64-bit edition of Office installed.

.PARAMETER ProductId
Required. ID must be set to a valid ProductRelease ID.
See https://support.microsoft.com/en-us/kb/2842297 for valid ids.

.PARAMETER LanguageId
Possible values match 'll-cc' pattern (Microsoft Language ids)
The ID value can be set to a valid Office culture language (such as en-us 
for English US or ja-jp for Japanese). The ll-cc value is the language 
identifier.
Defaults to the language from Get-Culture

.PARAMETER TargetFilePath
Full file path for the file to be output to.

.Example
New-ODTConfiguration -Bitness "64" -ProductId "O365ProPlusRetail" -TargetFilePath "$env:Public/Documents/config.xml"
Creates a config.xml file in public documents for installing the 64bit 
Office 365 ProPlus and sets the language to match the value in Get-Culture

.Example
New-ODTConfiguration -Bitness "64" -ProductId "O365ProPlusRetail" -TargetFilePath "$env:Public/Documents/config.xml" -LanguageId "es-es"
Creates a config.xml file in public documents for installing the 64bit 
Office 365 ProPlus and sets the language to Spanish

.Notes
Here is what the configuration file looks like when created from this function:

<Configuration>
  <Add OfficeClientEdition="64">
    <Product ID="O365ProPlusRetail">
      <Language ID="en-US" />
    </Product>
  </Add>
</Configuration>

#>
    Param(

    [Parameter(Mandatory=$true)]
    [string] $Bitness,

    [Parameter(Mandatory=$true, HelpMessage="Example: O365ProPlusRetail")]
    [string] $ProductId,

    [Parameter()]
    [string] $LanguageId = (Get-Culture | %{$_.Name}),

    [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
    [string] $TargetFilePath

    )

    Process{
        if (!$validProductIds.Contains($ProductId)) {
           throw "Invalid or Unsupported Product Id"
        }
        
        $pathSplit = Split-Path -Path $TargetFilePath
        $createDir = [system.io.directory]::CreateDirectory($pathSplit)

        #Create Document and Add root Configuration Element
        [System.XML.XMLDocument]$ConfigFile = New-Object System.XML.XMLDocument
        [System.XML.XMLElement]$ConfigurationRoot=$ConfigFile.CreateElement("Configuration")
        $ConfigFile.appendChild($ConfigurationRoot) | Out-Null

        #Add the Add Element under Configuration and set the Bitness
        [System.XML.XMLElement]$AddElement=$ConfigFile.CreateElement("Add")
        $ConfigurationRoot.appendChild($AddElement) | Out-Null
        $AddElement.SetAttribute("OfficeClientEdition",$Bitness) | Out-Null

        #Add the Product Element under Add and set the ID
        [System.XML.XMLElement]$ProductElement=$ConfigFile.CreateElement("Product")
        $AddElement.appendChild($ProductElement) | Out-Null
        $ProductElement.SetAttribute("ID",$ProductId) | Out-Null

        #Add the Language Element under Product and set the ID
        [System.XML.XMLElement]$LanguageElement=$ConfigFile.CreateElement("Language")
        $ProductElement.appendChild($LanguageElement) | Out-Null
        $LanguageElement.SetAttribute("ID",$LanguageId) | Out-Null
        $ConfigFile.Save($TargetFilePath) | Out-Null

        Write-Host
        Write-Host "The Office XML Configuration file has been saved to: $TargetFilePath"
         
    }
}

Function Add-ODTProduct{
<#
.SYNOPSIS
Modifies an existing configuration xml file to remove all or particular
click to run products.

.PARAMETER ExcludeApps
Array of IDs of Apps to exclude from install

.PARAMETER ProductId
Required. ID must be set to a valid ProductRelease ID.
See https://support.microsoft.com/en-us/kb/2842297 for valid ids.

.PARAMETER LanguageIds
Possible values match 'll-cc' pattern (Microsoft Language ids)
The ID value can be set to a valid Office culture language (such as en-us 
for English US or ja-jp for Japanese). The ll-cc value is the language 
identifier.

.PARAMETER ConfigPath
Full file path for the file to be modified and be output to.

.Example
Add-ODTProduct -ProductId "O365ProPlusRetail" -LanguageId ("en-US", "es-es") -ConfigPath "$env:Public/Documents/config.xml" -ExcludeApps ("Access", "InfoPath")
Sets config to add the English and Spanish version of office 365 ProPlus
excluding Access and InfoPath

.Example
Add-ODTProduct -ProductId "O365ProPlusRetail" -LanguageId ("en-US", "es-es) -ConfigPath "$env:Public/Documents/config.xml"
Sets config to add the English and Spanish version of office 365 ProPlus

.Notes
Here is what the portion of configuration file looks like when modified by this function:

<Configuration>
  <Add OfficeClientEdition="64" >
    <Product ID="O365ProPlusRetail">
      <Language ID="en-US" />
      <Language ID="es-es" />
      <ExcludeApp ID="Access">
      <ExcludeApp ID="InfoPath">
    </Product>
  </Add>
  ...
</Configuration>

#>
    Param(

        [Parameter(Mandatory=$true)]
        [string] $ProductId,

        [Parameter(Mandatory=$true)]
        [string[]] $LanguageIds,

        [Parameter()]
        [string[]] $ExcludeApps,

        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [string] $TargetFilePath

    )

    Process{
        #Load the file
        [System.XML.XMLDocument]$ConfigFile = New-Object System.XML.XMLDocument
        $ConfigFile.Load($TargetFilePath) | Out-Null

        #Check that the file is properly formatted
        if($ConfigFile.Configuration -eq $null){
            throw $NoConfigurationElement
        }

        if($ConfigFile.Configuration.Add -eq $null){
            throw $NoAddElement
        }

        #Set the desired values
        [System.XML.XMLElement]$ProductElement = $ConfigFile.Configuration.Add.Product | ?  ID -eq $ProductId
        if($ProductElement -eq $null){
            [System.XML.XMLElement]$ProductElement=$ConfigFile.CreateElement("Product")
            $ConfigFile.Configuration.Add.appendChild($ProductElement) | Out-Null
            $ProductElement.SetAttribute("Id", $ProductId) | Out-Null
        }


        foreach($LanguageId in $LanguageIds){
            [System.XML.XMLElement]$LanguageElement = $ProductElement.Language | ?  ID -eq $LanguageId
            if($LanguageElement -eq $null){
                [System.XML.XMLElement]$LanguageElement=$ConfigFile.CreateElement("Language")
                $ProductElement.appendChild($LanguageElement) | Out-Null
                $LanguageElement.SetAttribute("ID", $LanguageId) | Out-Null
            }
        }

        foreach($ExcludeApp in $ExcludeApps){
            [System.XML.XMLElement]$ExcludeAppElement = $ProductElement.ExcludeApp | ?  ID -eq $ExcludeApp
            if($ExcludeAppElement -eq $null){
                [System.XML.XMLElement]$ExcludeAppElement=$ConfigFile.CreateElement("ExcludeApp")
                $ProductElement.appendChild($ExcludeAppElement) | Out-Null
                $ExcludeAppElement.SetAttribute("ID", $ExcludeApp) | Out-Null
            }
        }

        $ConfigFile.Save($TargetFilePath) | Out-Null

        Write-Host
        Write-Host "The Office XML Configuration file has been saved to: $TargetFilePath"
    }

}

Function Remove-ODTProduct{
<#
.SYNOPSIS
Modifies an existing configuration xml file to remove all or particular
click to run products.

.PARAMETER All
Set this switch to remove all click to run products

.PARAMETER ProductId
Required. ID must be set to a valid ProductRelease ID.
See https://support.microsoft.com/en-us/kb/2842297 for valid ids.

.PARAMETER LanguageIds
Possible values match 'll-cc' pattern (Microsoft Language ids)
The ID value can be set to a valid Office culture language (such as en-us 
for English US or ja-jp for Japanese). The ll-cc value is the language 
identifier.

.PARAMETER ConfigPath
Full file path for the file to be modified and be output to.

.Example
Remove-ODTProduct -All -ConfigPath "$env:Public/Documents/config.xml"
Sets config to remove all click to run products

.Example
Remove-ODTProduct -ProductId "O365ProPlusRetail" -LanguageId "en-US" -ConfigPath "$env:Public/Documents/config.xml"
Sets config to remove the english version of office 365 ProPlus

.Notes
Here is what the portion of configuration file looks like when modified by this function:

<Configuration>
...
  <Remove>
    <Product ID="O365ProPlusRetail">
        <Language ID="en-US"
    </Product>
  </Remove>
</Configuration>

-or-

<Configuration>
...
  <Remove All="TRUE" />
</Configuration>

#>
    Param(

        [Parameter()]
        [switch] $All,

        [Parameter()]
        [string] $ProductId,

        [Parameter()]
        [string[]] $LanguageIds,

        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [string] $TargetFilePath

    )

    Process{
        #Load file from path
        [System.XML.XMLDocument]$ConfigFile = New-Object System.XML.XMLDocument
        $ConfigFile.Load($TargetFilePath) | Out-Null

        #Check to see if it has the proper root element
        if($ConfigFile.Configuration -eq $null){
            throw $NoConfigurationElement
        }

        #Get the Remove element if it exists
        [System.XML.XMLElement]$RemoveElement = $ConfigFile.Configuration.GetElementsByTagName("Remove").Item(0)
        if($ConfigFile.Configuration.Remove -eq $null){
            [System.XML.XMLElement]$RemoveElement=$ConfigFile.CreateElement("Remove")
            $ConfigFile.Configuration.appendChild($RemoveElement) | Out-Null
        }

        #Set the desired values
        if($All){
             $RemoveElement.SetAttribute("All", "True") | Out-Null
        }else{
            [System.XML.XMLElement]$ProductElement = $RemoveElement.Product | ?  ID -eq $ProductId
            if($ProductElement -eq $null){
                [System.XML.XMLElement]$ProductElement=$ConfigFile.CreateElement("Product")
                $RemoveElement.appendChild($ProductElement) | Out-Null
                $ProductElement.SetAttribute("ID", $ProductId) | Out-Null
            }
            foreach($LanguageId in $LanguageIds){
                [System.XML.XMLElement]$LanguageElement = $ProductElement.Language | ?  ID -eq $LanguageId
                if($LanguageElement -eq $null){
                    [System.XML.XMLElement]$LanguageElement=$ConfigFile.CreateElement("Language")
                    $ProductElement.appendChild($LanguageElement) | Out-Null
                    $LanguageElement.SetAttribute("ID", $LanguageId) | Out-Null
                }
            }
        }

        #Save the file
        $ConfigFile.Save($TargetFilePath) | Out-Null

        Write-Host
        Write-Host "The Office XML Configuration file has been saved to: $TargetFilePath"
    }

}

Function Set-ODTUpdates{
<#
.SYNOPSIS
Modifies an existing configuration xml file to enable/disable updates

.PARAMETER Enabled
Optional. If Enabled is set to TRUE, the Click-to-Run update system will 
check for updates. If it is set to FALSE, the Click-to-Run update system 
is dormant.

.PARAMETER UpdatePath
Optional. If UpdatePath is not set, Click-to-Run installations obtain updates 
from the Microsoft Click-to-Run source (Content Delivery Network or CDN). This is by default.
UpdatePath can specify a network, local, or HTTP path of a Click-to-Run source.
Environment variables can be used for network or local paths.

.PARAMETER TargetVersion
Optional. If TargetVersion is not set, Click-to-Run updates to the most 
recent version from the Microsoft Click-to-Run source. If TargetVersion 
is set to empty (""), Click-to-Run updates to the latest version from the 
Microsoft Click-to-Run source. TargetVersion can be set to an Office build number,
for example, 15.1.2.3. When the version is set, Office attempts to transition to
the specified version in the next update cycle.

.PARAMETER Deadline
Optional. Sets a deadline by when updates to Office must be applied. 
The deadline is specified in Coordinated Universal Time (UTC).
You can use Deadline with Target Version to make sure that Office is 
updated to a particular version by a particular date. We recommend that 
you set the deadline at least a week in the future to allow users time 
to install the updates.

.PARAMETER ConfigPath
Full file path for the file to be modified and be output to.

.Example
Set-ODTUpdates -Enabled "False" -ConfigPath "$env:Public/Documents/config.xml"
Sets config to disable updates

.Example
Set-ODTUpdates -Enabled "True" -UpdatePath "\\Server\share\" -ConfigPath "$env:Public/Documents/config.xml" -Deadline "05/16/2014 18:30" -TargetVersion "15.1.2.3"
Office updates are enabled, update path is \\Server\share\, the product 
version is set to 15.1.2.3, and the deadline is set to May 16, 2014 at 6:30 PM UTC.

.Notes
Here is what the portion of configuration file looks like when modified by this function:

<Configuration>
  ...
  <Updates Enabled="TRUE" UpdatePath="\\Server\share\" TargetVersion="15.1.2.3" Deadline="05/16/2014 18:30"/>
  ...
</Configuration>

#>
    Param(

        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [string] $TargetFilePath,
        
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
        #Load the file
        [System.XML.XMLDocument]$ConfigFile = New-Object System.XML.XMLDocument
        $ConfigFile.Load($TargetFilePath) | Out-Null

        #Check to make sure the correct root element exists
        if($ConfigFile.Configuration -eq $null){
            throw $NoConfigurationElement
        }

        #Get the Updates Element if it exists
        [System.XML.XMLElement]$UpdateElement = $ConfigFile.Configuration.GetElementsByTagName("Updates").Item(0)
        if($ConfigFile.Configuration.Updates -eq $null){
            [System.XML.XMLElement]$UpdateElement=$ConfigFile.CreateElement("Updates")
            $ConfigFile.Configuration.appendChild($UpdateElement) | Out-Null
        }

        #Set the desired values
        if([string]::IsNullOrWhiteSpace($Enabled) -eq $false){
            $UpdateElement.SetAttribute("Enabled", $Enabled) | Out-Null
        }
        if([string]::IsNullOrWhiteSpace($UpdatePath) -eq $false){
            $UpdateElement.SetAttribute("UpdatePath", $UpdatePath) | Out-Null
        }
        if([string]::IsNullOrWhiteSpace($TargetVersion) -eq $false){
            $UpdateElement.SetAttribute("TargetVersion", $TargetVersion) | Out-Null
        }
        if([string]::IsNullOrWhiteSpace($Deadline) -eq $false){
            $UpdateElement.SetAttribute("Deadline", $Deadline) | Out-Null
        }

        $ConfigFile.Save($TargetFilePath) | Out-Null

        return $TargetFilePath
    }
}

Function Set-ODTConfigProperties{
<#
.SYNOPSIS
Modifies an existing configuration xml file to enable/disable updates

.PARAMETER AutoActivate
If AUTOACTIVATE is set to 1, the specified products will attempt to activate automatically. 
If AUTOACTIVATE is not set, the user may see the Activation Wizard UI.
You must not set AUTOACTIVATE for Office 365 Click-to-Run products. 

.PARAMETER ForceAppShutDown
An installation or removal action may be blocked if Office applications are running. 
Normally, such cases would start a process killer UI. Administrators can set 
FORCEAPPSHUTDOWN value to TRUE to prevent dependence on user interaction. When 
FORCEAPPSHUTDOWN is set to TRUE, any applications that block the action will be shut 
down. Data loss may occur. When FORCEAPPSHUTDOWN is set to FALSE (default), the 
action may fail if Office applications are running.

.PARAMETER PackageGUID
Optional. By default, all Office 2013 App-V packages created by using the Office 
Deployment Tool share the same App-V Package ID. Administrators can use PACKAGEGUID 
to specify a different Package ID. Also, PACKAGEGUID needs to be at least 25 
characters in length and be separated into 5 sections, with each section separated by 
a dash. The sections need to have the following number of characters: 8, 4, 4, 4, and 12. 

.PARAMETER SharedComputerLicensing
Optional. Set SharedComputerLicensing to 1 if you deploy Office 365 ProPlus to shared 
computers by using Remote Desktop Services.

.PARAMETER ConfigPath
Full file path for the file to be modified and be output to.

.Example
Set-ODTConfigProperties -AutoActivate "1" -ConfigPath "$env:Public/Documents/config.xml"
Sets config to automatically activate the products

.Example
Set-ODTConfigProperties -ForceAppShutDown "True" -PackageGUID "12345678-ABCD-1234-ABCD-1234567890AB" -ConfigPath "$env:Public/Documents/config.xml"
Sets the config so that apps are forced to shutdown during install and the package guid
to "12345678-ABCD-1234-ABCD-1234567890AB"

.Notes
Here is what the portion of configuration file looks like when modified by this function:

<Configuration>
  ...
  <Property Name="AUTOACTIVATE" Value="1" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
  <Property Name="PACKAGEGUID" Value="12345678-ABCD-1234-ABCD-1234567890AB" />
  <Property Name="SharedComputerLicensing" Value="0" />
  ...
</Configuration>

#>
    Param(
        
        [Parameter()]
        [string] $AutoActivate,

        [Parameter()]
        [string] $ForceAppShutDown,

        [Parameter()]
        [string] $PackageGUID,

        [Parameter()]
        [string] $SharedComputerLicensing,

        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [string] $TargetFilePath
    )

    Process{
        #Load file
        [System.XML.XMLDocument]$ConfigFile = New-Object System.XML.XMLDocument
        $ConfigFile.Load($TargetFilePath) | Out-Null

        #Check for proper root element
        if($ConfigFile.Configuration -eq $null){
            throw $NoConfigurationElement
        }

        #Set each property as desired
        if([string]::IsNullOrWhiteSpace($AutoActivate) -eq $false){
            [System.XML.XMLElement]$AutoActivateElement = $ConfigFile.Configuration.Property | ?  Name -eq "AUTOACTIVATE"
            if($AutoActivateElement -eq $null){
                [System.XML.XMLElement]$AutoActivateElement=$ConfigFile.CreateElement("Property")
            }
                
            $ConfigFile.Configuration.appendChild($AutoActivateElement) | Out-Null
            $AutoActivateElement.SetAttribute("Name", "AUTOACTIVATE") | Out-Null
            $AutoActivateElement.SetAttribute("Value", $AutoActivate) | Out-Null
        }

        if([string]::IsNullOrWhiteSpace($ForceAppShutDown) -eq $false){
            [System.XML.XMLElement]$ForceAppShutDownElement = $ConfigFile.Configuration.Property | ?  Name -eq "FORCEAPPSHUTDOWN"
            if($ForceAppShutDownElement -eq $null){
                [System.XML.XMLElement]$ForceAppShutDownElement=$ConfigFile.CreateElement("Property")
            }
                
            $ConfigFile.Configuration.appendChild($ForceAppShutDownElement) | Out-Null
            $ForceAppShutDownElement.SetAttribute("Name", "FORCEAPPSHUTDOWN") | Out-Null
            $ForceAppShutDownElement.SetAttribute("Value", $ForceAppShutDownElement) | Out-Null
        }

        if([string]::IsNullOrWhiteSpace($PackageGUID) -eq $false){
            [System.XML.XMLElement]$PackageGUIDElement = $ConfigFile.Configuration.Property | ?  Name -eq "PACKAGEGUID"
            if($PackageGUIDElement -eq $null){
                [System.XML.XMLElement]$PackageGUIDElement=$ConfigFile.CreateElement("Property")
            }
                
            $ConfigFile.Configuration.appendChild($PackageGUIDElement) | Out-Null
            $PackageGUIDElement.SetAttribute("Name", "PACKAGEGUID") | Out-Null
            $PackageGUIDElement.SetAttribute("Value", $PackageGUID) | Out-Null
        }

        if([string]::IsNullOrWhiteSpace($SharedComputerLicensing) -eq $false){
            [System.XML.XMLElement]$SharedComputerLicensingElement = $ConfigFile.Configuration.Property | ?  Name -eq "SharedComputerLicensing"
            if($SharedComputerLicensingElement -eq $null){
                [System.XML.XMLElement]$SharedComputerLicensingElement=$ConfigFile.CreateElement("Property")
            }
                
            $ConfigFile.Configuration.appendChild($SharedComputerLicensingElement) | Out-Null
            $SharedComputerLicensingElement.SetAttribute("Name", "SharedComputerLicensing") | Out-Null
            $SharedComputerLicensingElement.SetAttribute("Value", $SharedComputerLicensing) | Out-Null
        }

        $ConfigFile.Save($TargetFilePath) | Out-Null
        return $TargetFilePath
        
    }
}

Function Set-ODTAdd{
<#
.SYNOPSIS
Modifies an existing configuration xml file to enable/disable updates

.PARAMETER SourcePath
Optional.
The SourcePath value can be set to a network, local, or HTTP path that contains a 
Click-to-Run source. Environment variables can be used for network or local paths.
SourcePath indicates the location to save the Click-to-Run installation source 
when you run the Office Deployment Tool in download mode.
SourcePath indicates the installation source path from which to install Office 
when you run the Office Deployment Tool in configure mode. If you don’t specify 
SourcePath in configure mode, Setup will look in the current folder for the Office 
source files. If the Office source files aren’t found in the current folder, Setup 
will look on Office 365 for them.
SourcePath specifies the path of the Click-to-Run Office source from which the 
App-V package will be made when you run the Office Deployment Tool in packager mode.
If you do not specify SourcePath, Setup will attempt to create an \Office\Data\... 
folder structure in the working directory from which you are running setup.exe.

.PARAMETER Version
Optional. If a Version value is not set, the Click-to-Run product installation streams 
the latest available version from the source. The default is to use the most recently 
advertised build (as defined in v32.CAB or v64.CAB at the Click-to-Run Office installation source).
Version can be set to an Office 2013 build number by using this format: X.X.X.X

.PARAMETER Bitness
Required. Specifies the edition of Click-to-Run for Office 365 product to use: 32- or 64-bit.

.PARAMETER ConfigPath
Full file path for the file to be modified and be output to.

.Example
Set-ODTAdd -SourcePath "C:\Preload\Office" -ConfigPath "$env:Public/Documents/config.xml"
Sets config SourcePath property of the add element to C:\Preload\Office

.Example
Set-ODTAdd -SourcePath "C:\Preload\Office" -Version "15.1.2.3" -ConfigPath "$env:Public/Documents/config.xml"
Sets config SourcePath property of the add element to C:\Preload\Office and version to 15.1.2.3

.Notes
Here is what the portion of configuration file looks like when modified by this function:

<Configuration>
  ...
  <Add SourcePath="\\server\share\" Version="15.1.2.3" OfficeClientEdition="32"> 
      ...
  </Add>
  ...
</Configuration>

#>
    Param(

        [Parameter()]
        [string] $SourcePath,

        [Parameter()]
        [string] $Version,

        [Parameter()]
        [string] $Bitness,

        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [string] $TargetFilePath

    )

    Process{
        #Load file
        [System.XML.XMLDocument]$ConfigFile = New-Object System.XML.XMLDocument
        $ConfigFile.Load($TargetFilePath) | Out-Null

        #Check for proper root element
        if($ConfigFile.Configuration -eq $null){
            throw $NoConfigurationElement
        }

        #Get Add element if it exists
        if($ConfigFile.Configuration.Add -eq $null){
            [System.XML.XMLElement]$AddElement=$ConfigFile.CreateElement("Add")
            $ConfigFile.Configuration.appendChild($AddElement) | Out-Null
        }

        #Set values as desired
        if([string]::IsNullOrWhiteSpace($SourcePath) -eq $false){
            $ConfigFile.Configuration.Add.SetAttribute("SourcePath", $SourcePath) | Out-Null
        }
        if([string]::IsNullOrWhiteSpace($Version) -eq $false){
            $ConfigFile.Configuration.Add.SetAttribute("Version", $Version) | Out-Null
        }
        if([string]::IsNullOrWhiteSpace($Bitness) -eq $false){
            $ConfigFile.Configuration.Add.SetAttribute("OfficeClientEdition", $Bitness) | Out-Null
        }

        $ConfigFile.Save($TargetFilePath) | Out-Null

        return $TargetFilePath

    }

}

Function Set-ODTLogging{
<#
.SYNOPSIS
Modifies an existing configuration xml file to enable/disable updates

.PARAMETER Level
Optional. Specifies options for the logging that Click-to-Run Setup 
performs. The default level is Standard.

.PARAMETER Path
Optional. Specifies the fully qualified path of the folder that is 
used for the log file. You can use environment variables. The default is %temp%.

.PARAMETER ConfigPath
Full file path for the file to be modified and be output to.

.Example
Set-ODTLogging -Level "Off" -ConfigPath "$env:Public/Documents/config.xml"
Sets config to turn off logging

.Example
Set-ODTLogging -Level "Standard" -Path "%temp%" -ConfigPath "$env:Public/Documents/config.xml"
Sets config to turn logging on and store the logs in the temp folder

.Notes
Here is what the portion of configuration file looks like when modified by this function:

<Configuration>
  ...
  <Logging Level="Standard" Path="%temp%" />
  ...
</Configuration>

#>
    Param(

        [Parameter()]
        [string] $Level,

        [Parameter()]
        [string] $Path,

        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [string] $TargetFilePath

    )

    Process{
        #Load file
        [System.XML.XMLDocument]$ConfigFile = New-Object System.XML.XMLDocument
        $ConfigFile.Load($TargetFilePath) | Out-Null

        #Check for proper root element
        if($ConfigFile.Configuration -eq $null){
            throw $NoConfigurationElement
        }

        #Get logging element if it exists
        [System.XML.XMLElement]$LoggingElement = $ConfigFile.Configuration.GetElementsByTagName("Logging").Item(0)
        if($ConfigFile.Configuration.Logging -eq $null){
            [System.XML.XMLElement]$LoggingElement=$ConfigFile.CreateElement("Logging")
            $ConfigFile.Configuration.appendChild($LoggingElement) | Out-Null
        }

        #Set values
        if([string]::IsNullOrWhiteSpace($Level) -eq $false){
            $LoggingElement.SetAttribute("Level", $Level) | Out-Null
        }
        if([string]::IsNullOrWhiteSpace($Path) -eq $false){
            $LoggingElement.SetAttribute("Path", $Path) | Out-Null
        }

        $ConfigFile.Save($TargetFilePath) | Out-Null

        return $TargetFilePath

    }
}

Function Set-ODTDisplay{
<#
.SYNOPSIS
Modifies an existing configuration xml file to enable/disable updates

.PARAMETER Level
Optional. Determines the user interface that the user sees when the 
operation is performed. If Level is set to None, the user sees no UI. 
No progress UI, completion screen, error dialog boxes, or first run 
automatic start UI are displayed. If Level is set to Full, the user 
sees the normal Click-to-Run user interface: Automatic start, 
application splash screen, and error dialog boxes.

.PARAMETER AcceptEULA
If this attribute is set to TRUE, the user does not see a Microsoft 
Software License Terms dialog box. If this attribute is set to FALSE 
or is not set, the user may see a Microsoft Software License Terms dialog box.

.PARAMETER ConfigPath
Full file path for the file to be modified and be output to.

.Example
Set-ODTLogging -Level "Full" -ConfigPath "$env:Public/Documents/config.xml"
Sets config show the UI during install

.Example
Set-ODTDisplay -Level "none" -AcceptEULA "True" -ConfigPath "$env:Public/Documents/config.xml"
Sets config to hide UI and automatically accept EULA during install

.Notes
Here is what the portion of configuration file looks like when modified by this function:

<Configuration>
  ...
  <Display Level="None" AcceptEULA="TRUE" />
  ...
</Configuration>

#>
    Param(

        [Parameter()]
        [string] $Level,

        [Parameter()]
        [string] $AcceptEULA,

        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [string] $TargetFilePath

    )

    Process{
        #Load file
        [System.XML.XMLDocument]$ConfigFile = New-Object System.XML.XMLDocument
        $ConfigFile.Load($TargetFilePath) | Out-Null

        #Check for proper root element
        if($ConfigFile.Configuration -eq $null){
            throw $NoConfigurationElement
        }

        #Get display element if it exists
        [System.XML.XMLElement]$DisplayElement = $ConfigFile.Configuration.GetElementsByTagName("Display").Item(0)
        if($ConfigFile.Configuration.Display -eq $null){
            [System.XML.XMLElement]$DisplayElement=$ConfigFile.CreateElement("Display")
            $ConfigFile.Configuration.appendChild($DisplayElement) | Out-Null
        }

        #Set values
        if([string]::IsNullOrWhiteSpace($Level) -eq $false){
            $DisplayElement.SetAttribute("Level", $Level) | Out-Null
        }
        if([string]::IsNullOrWhiteSpace($Path) -eq $AcceptEULA){
            $DisplayElement.SetAttribute("AcceptEULA", $AcceptEULA) | Out-Null
        }

        $ConfigFile.Save($TargetFilePath) | Out-Null

        return $TargetFilePath

    }

}