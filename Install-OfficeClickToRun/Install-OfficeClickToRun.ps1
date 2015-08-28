Add-Type -TypeDefinition @"
   public enum OfficeCTRVersion
   {
      Office2013
   }
"@

$enum = "
using System;
 
namespace Microsoft.Office
{
     [FlagsAttribute]
     public enum Products
     {
         Unknown = 0,
         O365ProPlusRetail = 1,
         O365BusinessRetail = 2,
         VisioProRetail = 4,
         ProjectProRetail = 8,
         SPDRetail = 16
     }
}
"
Add-Type -TypeDefinition $enum -Language CSharpVersion3

$enum2 = "
using System;
 
    [FlagsAttribute]
    public enum LogLevel
    {
        None=0,
        Full=1
    }
"
Add-Type -TypeDefinition $enum2 -Language CSharpVersion3

function Install-OfficeClickToRun {
    [CmdletBinding()]
    Param(
        [Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true, Position=0)]
        [string] $ConfigurationXML = $NULL,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $TargetFilePath = $NULL,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [OfficeCTRVersion] $OfficeVersion = "Office2013"
    )

    #Load the file
    [System.XML.XMLDocument]$ConfigFile = New-Object System.XML.XMLDocument
        
    if ($TargetFilePath) {
        $ConfigFile.Load($TargetFilePath) | Out-Null
    } else {
        if ($ConfigurationXml) 
        {
            $ConfigFile.LoadXml($ConfigurationXml) | Out-Null
            $global:saveLastConfigFile = $NULL
            $TargetFilePath = $NULL
        }
    }

    $officeCtrPath = Join-Path $PSScriptRoot "Office2013Setup.exe"

    if (!(Test-Path -Path $officeCtrPath)) {
       $officeCtrPath = Join-Path $PSScriptRoot "Setup.exe"
    }

    if ($OfficeVersion -eq "Office2013") {
        if (!(Test-Path -Path $officeCtrPath)) {
           throw "Cannot find the Office 2013 Setup executable"
        }
    }
    
    if (!($TargetFilePath)) {
      if ($ConfigurationXML) {
         $TargetFilePath = Join-Path $PSScriptRoot "configuration.xml"
         New-Item -Path $TargetFilePath -ItemType "File" -Value $ConfigurationXML -Force | Out-Null
      }
    }
    $products = Get-ODTProductToAdd -TargetFilePath $TargetFilePath 
    $addNode = Get-ODTAdd -TargetFilePath $TargetFilePath 

    $sourcePath = $addNode.SourcePath
    $version = $addNode.Version
    $edition = $addNode.OfficeClientEdition

    foreach ($product in $products)
    {
        $languages = getProductLanguages -Product $product 
        $existingLangs = checkForLanguagesInSourceFiles -Languages $languages -SourcePath $sourcePath -Version $version -Edition $edition
        Set-ODTProductToAdd -TargetFilePath $TargetFilePath -ProductId $product.ProductId -LanguageIds $existingLangs | Out-Null
    }

    Set-ODTDisplay -TargetFilePath $TargetFilePath -Level None -AcceptEULA $true | Out-Null

    $cmdLine = $officeCtrPath + " /configure " + $TargetFilePath

    Write-Host "Installing Office Click-To-Run..."
    Invoke-Expression -Command  $cmdLine
}

Function checkForLanguagesInSourceFiles() {
    [CmdletBinding()]
    Param(
        [Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true, Position=0)]
        $Languages = $NULL,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string]$SourcePath = $NULL,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string]$Version = $NULL,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string]$Edition = $NULL
    )

    $returnLanguages = @()

    if (!($SourcePath)) {
      $localSource = Join-Path $PSScriptRoot "Office\Data"
      if (Test-Path -Path $localSource) {
         $SourcePath = $PSScriptRoot
      }
    }

    if (!($Version)) {
       $localPath = $env:TEMP
       $cabPath = Join-Path $PSScriptRoot "Office\Data\v$Edition.cab"
       $cabFolderPath = Join-Path $PSScriptRoot "Office\Data"
       $vdXmlPath = Join-Path $localPath "\VersionDescriptor.xml"
       
       if (Test-Path -Path $cabPath) {
          Invoke-Expression -Command "Expand $cabPath -F:VersionDescriptor.xml $localPath" | Out-Null
          $Version = getVersionFromVersionDescriptor -vesionDescriptorPath $vdXmlPath
          Remove-Item -Path $vdXmlPath -Force
       }
    }

    $verionDir = Join-Path $PSScriptRoot "Office\Data\$Version"
    
    if (Test-Path -Path $verionDir) {
       foreach ($lang in $Languages) {
          $fileName = "stream.x86.$lang.dat"
          if ($Edition -eq "64") {
             $fileName = "stream.x64.$lang.dat"
          }
          
          $langFile = Join-Path $verionDir $fileName 
          
          if (Test-Path -Path $langFile) {
             $returnLanguages += $lang
          }
       }
    }

    return $returnLanguages
}

Function getVersionFromVersionDescriptor() {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [string] $vesionDescriptorPath = $NULL
    )

    [System.XML.XMLDocument]$doc = New-Object System.XML.XMLDocument

    if ($vesionDescriptorPath) {
        $doc.Load($vesionDescriptorPath) | Out-Null
        return $doc.DocumentElement.Available.Build
    }
}

Function getProductLanguages() {
    [CmdletBinding()]
    Param(
        [Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true, Position=0)]
        $Product = $NULL
    )

    $languages = @()

    foreach ($language in $Product.Languages)
    {
      if (!($languages.Contains($language))) {
          $languages += $language
      }
    }

    return $languages
}

Function getUniqueLanguages() {
    [CmdletBinding()]
    Param(
        [Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true, Position=0)]
        $Products = $NULL
    )

    $languages = @()
    foreach ($product in $Products)
    {
       foreach ($language in $product.Languages)
       {
          if (!($languages.Contains($language))) {
            $languages += $language
          }
       }
    }

    return $languages
}

Function Get-ODTProductToAdd{
<#
.SYNOPSIS
Gets list of Products and the corresponding language and exlcudeapp values
from the specified configuration file

.PARAMETER All
Switch to return All Products

.PARAMETER ProductId
Id of Product that you want to pull from the configuration file

.PARAMETER TargetFilePath
Required. Full file path for the file.

.Example
Get-ODTProductToAdd -All -TargetFilePath "$env:Public\Documents\config.xml"
Returns all Products and their corresponding Language and Exclude values
if they have them 

.Example
Get-ODTProductToAdd -ProductId "O365ProPlusRetail" -TargetFilePath "$env:Public\Documents\config.xml"
Returns the Product with the O365ProPlusRetail Id and its corresponding
Language and Exclude values

#>
    [CmdletBinding()]
    Param(
        [Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true, Position=0)]
        [string] $ConfigurationXML = $NULL,

        [Parameter(ParameterSetName="ID",Mandatory=$true)]
        [Microsoft.Office.Products] $ProductId = "Unknown",

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $TargetFilePath,

        [Parameter(ParameterSetName="All")]
        [switch] $All
    )

    Process{
        $TargetFilePath = GetFilePath -TargetFilePath $TargetFilePath

        #Load the file
        [System.XML.XMLDocument]$ConfigFile = New-Object System.XML.XMLDocument

        if ($TargetFilePath) {
           $ConfigFile.Load($TargetFilePath) | Out-Null
        } else {
            if ($ConfigurationXml) 
            {
              $ConfigFile.LoadXml($ConfigurationXml) | Out-Null
              $global:saveLastConfigFile = $NULL
              $global:saveLastFilePath = $NULL
            }
        }

        #Check that the file is properly formatted
        if($ConfigFile.Configuration -eq $null){
            throw $NoConfigurationElement
        }

        if($ConfigFile.Configuration.Add -eq $null){
            throw $NoAddElement
        }

        if($PSCmdlet.ParameterSetName -eq "All"){
            foreach($ProductElement in $ConfigFile.Configuration.Add.Product){
                $Result = New-Object –TypeName PSObject 

                Add-Member -InputObject $Result -MemberType NoteProperty -Name "ProductId" -Value ($ProductElement.GetAttribute("ID"))

                if($ProductElement.Language -ne $null){
                    Add-Member -InputObject $Result -MemberType NoteProperty -Name "Languages" -Value ($ProductElement.Language.GetAttribute("ID"))
                }

                if($ProductElement.ExcludeApp -ne $null){
                    Add-Member -InputObject $Result -MemberType NoteProperty -Name "ExcludedApps" -Value ($ProductElement.ExcludeApp.GetAttribute("ID"))
                }
                $Result
            }
        }else{
            [System.XML.XMLElement]$ProductElement = $ConfigFile.Configuration.Add.Product | ?  ID -eq $ProductId
            $Result = New-Object –TypeName PSObject 
            Add-Member -InputObject $Result -MemberType NoteProperty -Name "ProductId" -Value ($ProductElement.GetAttribute("ID"))
            if($ProductElement.Language -ne $null){
                Add-Member -InputObject $Result -MemberType NoteProperty -Name "Languages" -Value ($ProductElement.Language.GetAttribute("ID"))
            }

            if($ProductElement.ExcludeApp -ne $null){
                Add-Member -InputObject $Result -MemberType NoteProperty -Name "ExcludedApps" -Value ($ProductElement.ExcludeApp.GetAttribute("ID"))
            }
            $Result
        }

    }

}

Function Get-ODTAdd{
<#
.SYNOPSIS
Gets the value of the Add section in the configuration file

.PARAMETER TargetFilePath
Required. Full file path for the file.

.Example
Get-ODTAdd -TargetFilePath "$env:Public\Documents\config.xml"
Returns the value of the Add section if it exists in the specified
file. 

#>
    Param(

        [Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true, Position=0)]
        [string] $ConfigurationXML = $NULL,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $TargetFilePath

    )

    Process{
        $TargetFilePath = GetFilePath -TargetFilePath $TargetFilePath

        #Load the file
        [System.XML.XMLDocument]$ConfigFile = New-Object System.XML.XMLDocument

        if ($TargetFilePath) {
           $ConfigFile.Load($TargetFilePath) | Out-Null
        } else {
            if ($ConfigurationXml) 
            {
              $ConfigFile.LoadXml($ConfigurationXml) | Out-Null
              $global:saveLastConfigFile = $NULL
              $global:saveLastFilePath = $NULL
            }
        }

        #Check that the file is properly formatted
        if($ConfigFile.Configuration -eq $null){
            throw $NoConfigurationElement
        }
        
        $ConfigFile.Configuration.GetElementsByTagName("Add") | Select OfficeClientEdition, SourcePath, Version
    }

}

Function Set-ODTDisplay{
<#
.SYNOPSIS
Modifies an existing configuration xml file to set display level and acceptance of the EULA

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

.PARAMETER TargetFilePath
Full file path for the file to be modified and be output to.

.Example
Set-ODTLogging -Level "Full" -TargetFilePath "$env:Public/Documents/config.xml"
Sets config show the UI during install

.Example
Set-ODTDisplay -Level "none" -AcceptEULA "True" -TargetFilePath "$env:Public/Documents/config.xml"
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

        [Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true, Position=0)]
        [string] $ConfigurationXML = $NULL,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [LogLevel] $Level,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [bool] $AcceptEULA,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $TargetFilePath

    )

    Process{
        $TargetFilePath = GetFilePath -TargetFilePath $TargetFilePath

        #Load file
        [System.XML.XMLDocument]$ConfigFile = New-Object System.XML.XMLDocument

        if ($TargetFilePath) {
           $ConfigFile.Load($TargetFilePath) | Out-Null
        } else {
            if ($ConfigurationXml) 
            {
              $ConfigFile.LoadXml($ConfigurationXml) | Out-Null
              $global:saveLastConfigFile = $NULL
              $global:saveLastFilePath = $NULL
            }
        }

        $global:saveLastConfigFile = $ConfigFile.OuterXml

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
        } else {
            if ($PSBoundParameters.ContainsKey('Level')) {
                $ConfigFile.Configuration.Add.RemoveAttribute("Level")
            }
        }

        if([string]::IsNullOrWhiteSpace($Path) -eq $AcceptEULA){
            $DisplayElement.SetAttribute("AcceptEULA", $AcceptEULA) | Out-Null
        } else {
            if ($PSBoundParameters.ContainsKey('AcceptEULA')) {
                $ConfigFile.Configuration.Add.RemoveAttribute("AcceptEULA")
            }
        }

        $ConfigFile.Save($TargetFilePath) | Out-Null
        $global:saveLastFilePath = $TargetFilePath

        if (($PSCmdlet.MyInvocation.PipelineLength -eq 1) -or `
            ($PSCmdlet.MyInvocation.PipelineLength -eq $PSCmdlet.MyInvocation.PipelinePosition)) {
            Write-Host

            Format-XML ([xml](cat $TargetFilePath)) -indent 4

            Write-Host
            Write-Host "The Office XML Configuration file has been saved to: $TargetFilePath"
        } else {
            $results = new-object PSObject[] 0;
            $Result = New-Object –TypeName PSObject 
            Add-Member -InputObject $Result -MemberType NoteProperty -Name "TargetFilePath" -Value $TargetFilePath
            Add-Member -InputObject $Result -MemberType NoteProperty -Name "Level" -Value $Level
            Add-Member -InputObject $Result -MemberType NoteProperty -Name "AcceptEULA" -Value $AcceptEULA
            $Result
        }
    }

}

Function GetFilePath() {
    Param(
       [Parameter(ValueFromPipelineByPropertyName=$true)]
       [string] $TargetFilePath
    )

    if (!($TargetFilePath)) {
        $TargetFilePath = $global:saveLastFilePath
    }  

    if (!($TargetFilePath)) {
       Write-Host "Enter the path to the XML Configuration File: " -NoNewline
       $TargetFilePath = Read-Host
    } else {
       #Write-Host "Target XML Configuration File: $TargetFilePath"
    }

    return $TargetFilePath
}






