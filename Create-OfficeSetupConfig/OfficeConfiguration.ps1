[String]$global:saveLastConfigFile = $NULL
[String]$global:saveLastFilePath = $NULL

$validProductIds = @("O365ProPlusRetail","O365BusinessRetail","VisioProRetail","ProjectProRetail", "SPDRetail")

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

$validLanguages = @(
"English|en-us",
"Arabic|ar-sa",
"Bulgarian|bg-bg",
"Chinese (Simplified)|zh-cn",
"Chinese|zh-tw",
"Croatian|hr-hr",
"Czech|cs-cz",
"Danish|da-dk",
"Dutch|nl-nl",
"Estonian|et-ee",
"Finnish|fi-fi",
"French|fr-fr",
"German|de-de",
"Greek|el-gr",
"Hebrew|he-il",
"Hindi|hi-in",
"Hungarian|hu-hu",
"Indonesian|id-id",
"Italian|it-it",
"Japanese|ja-jp",
"Kazakh|kk-kh",
"Korean|ko-kr",
"Latvian|lv-lv",
"Lithuanian|lt-lt",
"Malay|ms-my",
"Norwegian (Bokmål)|nb-no",
"Polish|pl-pl",
"Portuguese|pt-br",
"Portuguese|pt-pt",
"Romanian|ro-ro",
"Russian|ru-ru",
"Serbian (Latin)|sr-latn-rs",
"Slovak|sk-sk",
"Slovenian|sl-si",
"Spanish|es-es",
"Swedish|sv-se",
"Thai|th-th",
"Turkish|tr-tr",
"Ukrainian|uk-ua")

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
    [CmdletBinding()]
    Param(

    [Parameter()]
    [string] $Bitness = $NULL,

    [Parameter(HelpMessage="Example: O365ProPlusRetail")]
    [Microsoft.Office.Products] $ProductId = "Unknown",

    [Parameter()]
    [string] $LanguageId = $NULL,

    [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
    [string] $TargetFilePath

    )

    Begin {
      $once = $false
    }

    Process{

        if ($ProductId -eq "Unknown") {
            $ProductId = SelectProductId
        }

        if (!$Bitness) {
            $Bitness = SelectBitness
        }

        $ProductId = IsValidProductId -ProductId $ProductId

        if (!($LanguageId)) {
            $LanguageId = (Get-Culture | %{$_.Name})
            $LanguageId = LanguagePrompt -DefaultLanguage $LanguageId
        }

        $LanguageId = IsSupportedLanguage -Language $LanguageId
        
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
        $global:saveLastFilePath = $TargetFilePath

        if ($PSCmdlet.MyInvocation.PipelineLength -eq 1) {
            Write-Host

            Format-XML ([xml](cat $TargetFilePath)) -indent 4

            Write-Host
            Write-Host "The Office XML Configuration file has been saved to: $TargetFilePath"
        } else {
            $results = new-object PSObject[] 0;
            $Result = New-Object –TypeName PSObject 
            Add-Member -InputObject $Result -MemberType NoteProperty -Name "TargetFilePath" -Value $TargetFilePath
            Add-Member -InputObject $Result -MemberType NoteProperty -Name "ProductId" -Value $ProductId
            Add-Member -InputObject $Result -MemberType NoteProperty -Name "LanguageId" -Value $LanguageId
            $Result
        }
         
    }
}

Function Undo-ODTLastChange {

    Process{
        if ($global:saveLastConfigFile -and $global:saveLastFilePath) {
            [System.XML.XMLDocument]$ConfigFile = New-Object System.XML.XMLDocument

            $ConfigFile.LoadXml($global:saveLastConfigFile) | Out-Null
            $ConfigFile.Save($global:saveLastFilePath) | Out-Null

            Write-Host

            Format-XML ([xml](cat $global:saveLastFilePath)) -indent 4

            Write-Host
            Write-Host "The Office XML Configuration file has been saved to: $global:saveLastFilePath"
        }
    }
}

Function Show-ODTConfiguration {
    Param(
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
        [string] $TargetFilePath
    )

    Process{        
        Write-Host

        Format-XML ([xml](cat $TargetFilePath)) -indent 4

        Write-Host
    }
}


Function Add-ODTProductToAdd{
<#
.SYNOPSIS
Modifies an existing configuration xml file to add a particular
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

.PARAMETER TargetFilePath
Full file path for the file to be modified and be output to.

.Example
Add-ODTProductToAdd -ProductId "O365ProPlusRetail" -LanguageId ("en-US", "es-es") -TargetFilePath "$env:Public/Documents/config.xml" -ExcludeApps ("Access", "InfoPath")
Sets config to add the English and Spanish version of office 365 ProPlus
excluding Access and InfoPath

.Example
Add-ODTProductToAdd -ProductId "O365ProPlusRetail" -LanguageId ("en-US", "es-es) -TargetFilePath "$env:Public/Documents/config.xml"
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
    [CmdletBinding()]
    Param(
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $TargetFilePath = $NULL,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [Microsoft.Office.Products] $ProductId = "Unknown",

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [Alias("LanguageId")]
        [string[]] $LanguageIds = @(),

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string[]] $ExcludeApps

    )

    Process{

        $TargetFilePath = GetFilePath -TargetFilePath $TargetFilePath

        if ($ProductId -eq "Unknown") {
           $ProductId = SelectProductId
        }

        $ProductId = IsValidProductId -ProductId $ProductId
        
        $langCount = $LanguageIds.Count

        if ($langCount -gt 0) {
           foreach ($language in $LanguageIds) {
              $language = IsSupportedLanguage -Language $language
           }
        } else {
            $CurrentLanguage = (Get-Culture | %{$_.Name})
            $LanguageIds += LanguagePrompt -DefaultLanguage $CurrentLanguage
        }

        #Load the file
        [System.XML.XMLDocument]$ConfigFile = New-Object System.XML.XMLDocument
        
        $ConfigFile.Load($TargetFilePath) | Out-Null

        $global:saveLastConfigFile = $ConfigFile.OuterXml

        #Check that the file is properly formatted
        if($ConfigFile.Configuration -eq $null){
            throw $NoConfigurationElement
        }

        [System.XML.XMLElement]$AddElement=$NULL
        if($ConfigFile.Configuration.Add -eq $null){
            $AddElement=$ConfigFile.CreateElement("Add")
            $ConfigFile.DocumentElement.appendChild($AddElement) | Out-Null
        } else {
           $AddElement = $ConfigFile.Configuration.Add 
        }

        #Set the desired values
        [System.XML.XMLElement]$ProductElement = $ConfigFile.Configuration.Add.Product | ?  ID -eq $ProductId
        if($ProductElement -eq $null){
            [System.XML.XMLElement]$ProductElement=$ConfigFile.CreateElement("Product")
            $AddElement.appendChild($ProductElement) | Out-Null
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

        foreach($ExcludeApp in $ExcludeApps){
            [System.XML.XMLElement]$ExcludeAppElement = $ProductElement.ExcludeApp | ?  ID -eq $ExcludeApp
            if($ExcludeAppElement -eq $null){
                [System.XML.XMLElement]$ExcludeAppElement=$ConfigFile.CreateElement("ExcludeApp")
                $ProductElement.appendChild($ExcludeAppElement) | Out-Null
                $ExcludeAppElement.SetAttribute("ID", $ExcludeApp) | Out-Null
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
            Add-Member -InputObject $Result -MemberType NoteProperty -Name "ProductId" -Value $ProductId
            $Result
        }
    }

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
        $ConfigFile.Load($TargetFilePath) | Out-Null

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

Function Remove-ODTProductToAdd{
<#
.SYNOPSIS
Removes an existing product to add from the configuration file

.PARAMETER ProductId
Required. ID must be set to a valid ProductRelease ID.
See https://support.microsoft.com/en-us/kb/2842297 for valid ids.

.PARAMETER TargetFilePath
Full file path for the file to be modified and be output to.

.Example
Remove-ODTProductToAdd -ProductId "O365ProPlusRetail" -TargetFilePath "$env:Public/Documents/config.xml"
Removes the ProductToAdd with the ProductId 'O365ProPlusRetail' from the XML Configuration file

</Configuration>

#>
    [CmdletBinding()]
    Param(
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $ProductId = "Unknown",

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $TargetFilePath,

        [Parameter(ParameterSetName="All", ValueFromPipelineByPropertyName=$true)]
        [switch] $All
    )

    Process{
        $TargetFilePath = GetFilePath -TargetFilePath $TargetFilePath

        if ($ProductId -eq "Unknown") {
            $ProductId = SelectProductId
        }

        $ProductId = IsValidProductId -ProductId $ProductId

        #Load the file
        [System.XML.XMLDocument]$ConfigFile = New-Object System.XML.XMLDocument
        $ConfigFile.Load($TargetFilePath) | Out-Null

        $global:saveLastConfigFile = $ConfigFile.OuterXml

        #Check that the file is properly formatted
        if($ConfigFile.Configuration -eq $null){
            throw $NoConfigurationElement
        }

        if($ConfigFile.Configuration.Add -eq $null){
            throw $NoAddElement
        }

        if ($All) {
            #Set the desired values
            [System.XML.XMLElement]$ProductElement = $ConfigFile.Configuration.Add.Product | ?  ID -eq $ProductId
            if($ProductElement -ne $null){
                $ConfigFile.Configuration.Add.removeChild($ProductElement) | Out-Null
            }

            if ($ConfigFile.Configuration.Add.Product.Count -eq 0) {
                [System.XML.XMLElement]$AddNode = $ConfigFile.SelectSingleNode("/Configuration/Add")
                if ($AddNode) {
                    $ConfigFile.Configuration.removeChild($AddNode) | Out-Null
                }
            }
        } else {
          $removeAll = $ConfigFile.Configuration.Product.RemoveAll()
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
            $Result
        }
    }

}


Function Add-ODTProductToRemove{
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

.PARAMETER TargetFilePath
Full file path for the file to be modified and be output to.

.Example
Add-ODTProductToRemove -All -TargetFilePath "$env:Public/Documents/config.xml"
Sets config to remove all click to run products

.Example
Add-ODTProductToRemove -ProductId "O365ProPlusRetail" -LanguageId "en-US" -TargetFilePath "$env:Public/Documents/config.xml"
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
    [CmdletBinding()]
    Param(

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $TargetFilePath,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [switch] $All,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [Microsoft.Office.Products] $ProductId = "Unknown",

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [Alias("LanguageId")]
        [string[]] $LanguageIds
    )

    Process{
        $TargetFilePath = GetFilePath -TargetFilePath $TargetFilePath

        if ($ProductId -eq "Unknown") {
           $ProductId = SelectProductId
        }

        $ProductId = IsValidProductId -ProductId $ProductId

        #Load file from path
        [System.XML.XMLDocument]$ConfigFile = New-Object System.XML.XMLDocument
        $ConfigFile.Load($TargetFilePath) | Out-Null

        $global:saveLastConfigFile = $ConfigFile.OuterXml

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
            Add-Member -InputObject $Result -MemberType NoteProperty -Name "ProductId" -Value $ProductId
            Add-Member -InputObject $Result -MemberType NoteProperty -Name "LanguageIds" -Value $LanguageIds
            $Result
        }
    }

}

Function Get-ODTProductToRemove{
<#
.SYNOPSIS
Gets list of Products and the corresponding language values
from the specified configuration file

.PARAMETER ProductId
Id of Product that you want to pull from the configuration file

.PARAMETER TargetFilePath
Required. Full file path for the file.

.Example
Get-ODTProductToRemove -TargetFilePath "$env:Public\Documents\config.xml"
Returns all Products and their corresponding Language and Exclude values
if they have them 

.Example
Get-ODTProductToRemove -ProductId "O365ProPlusRetail" -TargetFilePath "$env:Public\Documents\config.xml"
Returns the Product with the O365ProPlusRetail Id and its corresponding
Language and Exclude values

#>
    [cmdletbinding()]
    Param(

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [Microsoft.Office.Products] $ProductId = "Unknown",

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $TargetFilePath

    )

    Begin {
        $defaultDisplaySet = 'ProductId','Languages', 'ExcludedApps'

        $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet(‘DefaultDisplayPropertySet’,[string[]]$defaultDisplaySet)
        $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
    }

    Process{
        $TargetFilePath = GetFilePath -TargetFilePath $TargetFilePath
      
        #Load the file
        [System.XML.XMLDocument]$ConfigFile = New-Object System.XML.XMLDocument
        $ConfigFile.Load($TargetFilePath) | Out-Null

        #Check that the file is properly formatted
        if($ConfigFile.Configuration -eq $null){
            throw $NoConfigurationElement
        }

        if(!($ConfigFile.Configuration.Remove)){
            throw $NoAddElement
        }

        [System.XML.XMLElement[]]$ProductElements
        if ($ProductId -eq "Unknown") {
           $ProductElements = $ConfigFile.Configuration.Remove.Product
        } else {
           $ProductElements = $ConfigFile.Configuration.Remove.Product | Where {$_.ID -eq $ProductId}
        }

        $results = new-object PSObject[] 0;

        foreach($ProductElement in $ProductElements){
            $Result = New-Object –TypeName PSObject 
            Add-Member -InputObject $Result -MemberType NoteProperty -Name "ProductId" -Value ($ProductElement.GetAttribute("ID"))
            Add-Member -InputObject $Result -MemberType NoteProperty -Name "TargetFilePath" -Value $TargetFilePath 
            if($ProductElement.Language -ne $null){
                Add-Member -InputObject $Result -MemberType NoteProperty -Name "Languages" -Value ($ProductElement.Language.GetAttribute("ID"))
            }

            if($ProductElement.ExcludeApp -ne $null){
                Add-Member -InputObject $Result -MemberType NoteProperty -Name "ExcludedApps" -Value ($ProductElement.ExcludeApp.GetAttribute("ID"))
            }

            $Result | Add-Member MemberSet PSStandardMembers $PSStandardMembers

            $results += $Result
        }
        
        $results
    }

}

Function Remove-ODTProductToRemove{
<#
.SYNOPSIS
Removes an existing product to remove from the configuration file

.PARAMETER ProductId
Required. ID must be set to a valid ProductRelease ID.
See https://support.microsoft.com/en-us/kb/2842297 for valid ids.

.PARAMETER TargetFilePath
Full file path for the file to be modified and be output to.

.Example
Add-ODTProductToRemove -ProductId "O365ProPlusRetail" -TargetFilePath "$env:Public/Documents/config.xml"
Removes the ProductToRemove with the ProductId 'O365ProPlusRetail' from the XML Configuration file

</Configuration>

#>
    [cmdletbinding()]
    Param(
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [Microsoft.Office.Products] $ProductId = "Unknown",

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $TargetFilePath
    )

    Process{
        $TargetFilePath = GetFilePath -TargetFilePath $TargetFilePath

        if ($ProductId -eq "Unknown") {
           $ProductId = SelectProductId
        }

        #Load the file
        [System.XML.XMLDocument]$ConfigFile = New-Object System.XML.XMLDocument
        $ConfigFile.Load($TargetFilePath) | Out-Null

        $global:saveLastConfigFile = $ConfigFile.OuterXml

        #Check that the file is properly formatted
        if($ConfigFile.Configuration -eq $null){
            throw $NoConfigurationElement
        }

        if($ConfigFile.Configuration.Add -eq $null){
            throw $NoAddElement
        }

        #Set the desired values
        [System.XML.XMLElement]$ProductElement = $ConfigFile.Configuration.Remove.Product | ?  ID -eq $ProductId
        if($ProductElement -ne $null){
            $ConfigFile.Configuration.Remove.removeChild($ProductElement) | Out-Null
        }

        if ($ConfigFile.Configuration.Remove.Product.Count -eq 0) {
            [System.XML.XMLElement]$RemoveNode = $ConfigFile.SelectSingleNode("/Configuration/Remove")
            if ($RemoveNode) {
                $ConfigFile.Configuration.removeChild($RemoveNode) | Out-Null
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
            $Result
        }
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

.PARAMETER TargetFilePath
Full file path for the file to be modified and be output to.

.Example
Set-ODTUpdates -Enabled "False" -TargetFilePath "$env:Public/Documents/config.xml"
Sets config to disable updates

.Example
Set-ODTUpdates -Enabled "True" -UpdatePath "\\Server\share\" -TargetFilePath "$env:Public/Documents/config.xml" -Deadline "05/16/2014 18:30" -TargetVersion "15.1.2.3"
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
    [CmdletBinding()]
    Param(

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $TargetFilePath,
        
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $Enabled,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $UpdatePath,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $TargetVersion,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $Deadline

    )

    Process{
        $TargetFilePath = GetFilePath -TargetFilePath $TargetFilePath

        #Load the file
        [System.XML.XMLDocument]$ConfigFile = New-Object System.XML.XMLDocument
        $ConfigFile.Load($TargetFilePath) | Out-Null

        $global:saveLastConfigFile = $ConfigFile.OuterXml

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
        } else {
          if ($PSBoundParameters.ContainsKey('Enabled')) {
              $ConfigFile.Configuration.Updates.RemoveAttribute("Enabled")
          }
        }

        if([string]::IsNullOrWhiteSpace($UpdatePath) -eq $false){
            $UpdateElement.SetAttribute("UpdatePath", $UpdatePath) | Out-Null
        } else {
          if ($PSBoundParameters.ContainsKey('UpdatePath')) {
              $ConfigFile.Configuration.Updates.RemoveAttribute("UpdatePath")
          }
        }

        if([string]::IsNullOrWhiteSpace($TargetVersion) -eq $false){
            $UpdateElement.SetAttribute("TargetVersion", $TargetVersion) | Out-Null
        } else {
          if ($PSBoundParameters.ContainsKey('TargetVersion')) {
              $ConfigFile.Configuration.Updates.RemoveAttribute("TargetVersion")
          }
        }

        if([string]::IsNullOrWhiteSpace($Deadline) -eq $false){
            $UpdateElement.SetAttribute("Deadline", $Deadline) | Out-Null
        } else {
          if ($PSBoundParameters.ContainsKey('Deadline')) {
              $ConfigFile.Configuration.Updates.RemoveAttribute("Deadline")
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
            Add-Member -InputObject $Result -MemberType NoteProperty -Name "ProductId" -Value $ProductId
            Add-Member -InputObject $Result -MemberType NoteProperty -Name "Enabled" -Value $Enabled
            Add-Member -InputObject $Result -MemberType NoteProperty -Name "UpdatePath" -Value $UpdatePath
            Add-Member -InputObject $Result -MemberType NoteProperty -Name "TargetVersion" -Value $TargetVersion
            Add-Member -InputObject $Result -MemberType NoteProperty -Name "Deadline" -Value $Deadline
            $Result
        }
    }
}

Function Get-ODTUpdates{
<#
.SYNOPSIS
Gets the value of the Updates section in the configuration file

.PARAMETER TargetFilePath
Required. Full file path for the file.

.Example
Get-ODTUpdates -TargetFilePath "$env:Public\Documents\config.xml"
Returns the value of the Updates section if it exists in the specified
file. 

#>
    [CmdletBinding()]
    Param(

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $TargetFilePath

    )

    Process{
        $TargetFilePath = GetFilePath -TargetFilePath $TargetFilePath
     
        #Load the file
        [System.XML.XMLDocument]$ConfigFile = New-Object System.XML.XMLDocument
        $ConfigFile.Load($TargetFilePath) | Out-Null

        #Check that the file is properly formatted
        if($ConfigFile.Configuration -eq $null){
            throw $NoConfigurationElement
        }
        
        $ConfigFile.Configuration.GetElementsByTagName("Updates");
    }

}

Function Remove-ODTUpdates{
<#
.SYNOPSIS
Removes the update section from an existing configuration xml file

.PARAMETER TargetFilePath
Full file path for the file to be modified and be output to.

.Example
Set-ODTUpdates -TargetFilePath "$env:Public/Documents/config.xml"

.Notes
This is the section that would be removed when running this function

<Configuration>
  ...
  <Updates Enabled="TRUE" UpdatePath="\\Server\share\" TargetVersion="15.1.2.3" Deadline="05/16/2014 18:30"/>
  ...
</Configuration>

#>
    [CmdletBinding()]
    Param(
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $TargetFilePath
    )

    Process{
        $TargetFilePath = GetFilePath -TargetFilePath $TargetFilePath

        #Load the file
        [System.XML.XMLDocument]$ConfigFile = New-Object System.XML.XMLDocument
        $ConfigFile.Load($TargetFilePath) | Out-Null

        $global:saveLastConfigFile = $ConfigFile.OuterXml

        #Check to make sure the correct root element exists
        if($ConfigFile.Configuration -eq $null){
            throw $NoConfigurationElement
        }

        #Get the Updates Element if it exists
        [System.XML.XMLElement]$UpdateElement = $ConfigFile.Configuration.GetElementsByTagName("Updates").Item(0)
        if($ConfigFile.Configuration.Updates -ne $null){
            $ConfigFile.Configuration.removeChild($UpdateElement) | Out-Null
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
            $Result
        }
    }
}


Function Set-ODTConfigProperties{
<#
.SYNOPSIS
Modifies an existing configuration xml file to set property values

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

.PARAMETER TargetFilePath
Full file path for the file to be modified and be output to.

.Example
Set-ODTConfigProperties -AutoActivate "1" -TargetFilePath "$env:Public/Documents/config.xml"
Sets config to automatically activate the products

.Example
Set-ODTConfigProperties -ForceAppShutDown "True" -PackageGUID "12345678-ABCD-1234-ABCD-1234567890AB" -TargetFilePath "$env:Public/Documents/config.xml"
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
    [CmdletBinding()]
    Param(
        
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $AutoActivate,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $ForceAppShutDown,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $PackageGUID,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $SharedComputerLicensing,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $TargetFilePath
    )

    Process{
        $TargetFilePath = GetFilePath -TargetFilePath $TargetFilePath

        #Load file
        [System.XML.XMLDocument]$ConfigFile = New-Object System.XML.XMLDocument
        $ConfigFile.Load($TargetFilePath) | Out-Null

        $global:saveLastConfigFile = $ConfigFile.OuterXml

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
            $ForceAppShutDownElement.SetAttribute("Value", $ForceAppShutDown) | Out-Null
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
            Add-Member -InputObject $Result -MemberType NoteProperty -Name "SharedComputerLicensing" -Value $SharedComputerLicensing
            Add-Member -InputObject $Result -MemberType NoteProperty -Name "PackageGUID" -Value $PackageGUID
            Add-Member -InputObject $Result -MemberType NoteProperty -Name "ForceAppShutDown" -Value $ForceAppShutDown
            Add-Member -InputObject $Result -MemberType NoteProperty -Name "AutoActivate" -Value $AutoActivate
            $Result
        }
    }
}

Function Get-ODTConfigProperties{
<#
.SYNOPSIS
Gets the value of the ODTConfigProperties in the configuration file

.PARAMETER TargetFilePath
Required. Full file path for the file.

.Example
Get-ODTConfigProperties -TargetFilePath "$env:Public\Documents\config.xml"
Returns the value of the Properties if they exists in the specified
file. 

#>
    Param(

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $TargetFilePath

    )

    Process{
        $TargetFilePath = GetFilePath -TargetFilePath $TargetFilePath

        #Load the file
        [System.XML.XMLDocument]$ConfigFile = New-Object System.XML.XMLDocument
        $ConfigFile.Load($TargetFilePath) | Out-Null

        #Check that the file is properly formatted
        if($ConfigFile.Configuration -eq $null){
            throw $NoConfigurationElement
        }
        
        $ConfigFile.Configuration.GetElementsByTagName("Property")
    }

}

Function Remove-ODTConfigProperties{
<#
.SYNOPSIS
Removes the property items from an existing configuration xml file

.PARAMETER TargetFilePath
Full file path for the file to be modified and be output to.

.PARAMETER Name
Name of the property to remove

.Example
Remove-ODTConfigProperties -TargetFilePath "$env:Public/Documents/config.xml"
Removes all of the poperty items from the existing configuration xml file

.Example
Remove-ODTConfigProperties -Name "AUTOACTIVATE" -TargetFilePath "$env:Public/Documents/config.xml"
Removes the poperty items with the Name "AUTOACTIVATE" from the existing configuration xml file

.Notes
Here is what the portion of configuration file that would be removed by this function:

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
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $TargetFilePath,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $Name = $NULL
    )

    Process{
        $TargetFilePath = GetFilePath -TargetFilePath $TargetFilePath

        #Load file
        [System.XML.XMLDocument]$ConfigFile = New-Object System.XML.XMLDocument
        $ConfigFile.Load($TargetFilePath) | Out-Null

        $global:saveLastConfigFile = $ConfigFile.OuterXml

        #Check for proper root element
        if($ConfigFile.Configuration -eq $null){
            throw $NoConfigurationElement
        }

        if ($Name) {
          [System.XML.XMLElement]$ForceAppShutDownElement = $ConfigFile.Configuration.Property | ?  Name -eq $Name.ToUpper()
          if ($ForceAppShutDownElement) {
              $removeNode = $ConfigFile.Configuration.removeChild($ForceAppShutDownElement)
          }
        } else {
          $removeAll = $ConfigFile.Configuration.Property.RemoveAll()
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
            $Result
        }
    }
}


Function Set-ODTAdd{
<#
.SYNOPSIS
Modifies an existing configuration xml file's add section

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

.PARAMETER TargetFilePath
Full file path for the file to be modified and be output to.

.Example
Set-ODTAdd -SourcePath "C:\Preload\Office" -TargetFilePath "$env:Public/Documents/config.xml"
Sets config SourcePath property of the add element to C:\Preload\Office

.Example
Set-ODTAdd -SourcePath "C:\Preload\Office" -Version "15.1.2.3" -TargetFilePath "$env:Public/Documents/config.xml"
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

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $SourcePath = $NULL,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $Version,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $Bitness,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $TargetFilePath

    )

    Process{
        $TargetFilePath = GetFilePath -TargetFilePath $TargetFilePath

        #Load file
        [System.XML.XMLDocument]$ConfigFile = New-Object System.XML.XMLDocument
        $ConfigFile.Load($TargetFilePath) | Out-Null

        $global:saveLastConfigFile = $ConfigFile.OuterXml

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
        } else {
            if ($PSBoundParameters.ContainsKey('SourcePath')) {
                $ConfigFile.Configuration.Add.RemoveAttribute("SourcePath")
            }
        }

        if([string]::IsNullOrWhiteSpace($Version) -eq $false){
            $ConfigFile.Configuration.Add.SetAttribute("Version", $Version) | Out-Null
        } else {
            if ($PSBoundParameters.ContainsKey('Version')) {
                $ConfigFile.Configuration.Add.RemoveAttribute("Version")
            }
        }

        if([string]::IsNullOrWhiteSpace($Bitness) -eq $false){
            $ConfigFile.Configuration.Add.SetAttribute("OfficeClientEdition", $Bitness) | Out-Null
        } else {
            if ($PSBoundParameters.ContainsKey('OfficeClientEdition')) {
                $ConfigFile.Configuration.Add.RemoveAttribute("OfficeClientEdition")
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
            Add-Member -InputObject $Result -MemberType NoteProperty -Name "SourcePath" -Value $SourcePath
            Add-Member -InputObject $Result -MemberType NoteProperty -Name "Version" -Value $Version
            Add-Member -InputObject $Result -MemberType NoteProperty -Name "Bitness" -Value $Bitness
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

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $TargetFilePath

    )

    Process{
        $TargetFilePath = GetFilePath -TargetFilePath $TargetFilePath

        #Load the file
        [System.XML.XMLDocument]$ConfigFile = New-Object System.XML.XMLDocument
        $ConfigFile.Load($TargetFilePath) | Out-Null

        #Check that the file is properly formatted
        if($ConfigFile.Configuration -eq $null){
            throw $NoConfigurationElement
        }
        
        $ConfigFile.Configuration.GetElementsByTagName("Add") | Select OfficeClientEdition, SourcePath, Version
    }

}

Function Remove-ODTAdd{
<#
.SYNOPSIS
Removes the Add node from existing configuration xml file

.PARAMETER TargetFilePath
Full file path for the file to be modified and be output to.

.Example
Set-ODTAdd -SourcePath "C:\Preload\Office" -TargetFilePath "$env:Public/Documents/config.xml"
Sets config SourcePath property of the add element to C:\Preload\Office

.Example
Remove-ODTAdd -TargetFilePath "$env:Public/Documents/config.xml"
Removes the Add node from the xml congfiguration file

#>
    Param(
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $TargetFilePath
    )

    Process{
        $TargetFilePath = GetFilePath -TargetFilePath $TargetFilePath

        #Load file
        [System.XML.XMLDocument]$ConfigFile = New-Object System.XML.XMLDocument
        $ConfigFile.Load($TargetFilePath) | Out-Null

        $global:saveLastConfigFile = $ConfigFile.OuterXml

        #Check for proper root element
        if($ConfigFile.Configuration -eq $null){
            throw $NoConfigurationElement
        }

        $addNode = $ConfigFile.SelectSingleNode("/Configuration/Add")
        if ($addNode) {
            $removeAll = $ConfigFile.Configuration.removeChild($addNode)
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
            $Result
        }
    }

}


Function Set-ODTLogging{
<#
.SYNOPSIS
Modifies an existing configuration xml file to enable/disable logging

.PARAMETER Level
Optional. Specifies options for the logging that Click-to-Run Setup 
performs. The default level is Standard.

.PARAMETER Path
Optional. Specifies the fully qualified path of the folder that is 
used for the log file. You can use environment variables. The default is %temp%.

.PARAMETER TargetFilePath
Full file path for the file to be modified and be output to.

.Example
Set-ODTLogging -Level "Off" -TargetFilePath "$env:Public/Documents/config.xml"
Sets config to turn off logging

.Example
Set-ODTLogging -Level "Standard" -Path "%temp%" -TargetFilePath "$env:Public/Documents/config.xml"
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

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $Level,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $Path,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $TargetFilePath

    )

    Process{
        $TargetFilePath = GetFilePath -TargetFilePath $TargetFilePath

        #Load file
        [System.XML.XMLDocument]$ConfigFile = New-Object System.XML.XMLDocument
        $ConfigFile.Load($TargetFilePath) | Out-Null

        $global:saveLastConfigFile = $ConfigFile.OuterXml

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
        } else {
            if ($PSBoundParameters.ContainsKey('Level')) {
                $ConfigFile.Configuration.Add.RemoveAttribute("Level")
            }
        }

        if([string]::IsNullOrWhiteSpace($Path) -eq $false){
            $LoggingElement.SetAttribute("Path", $Path) | Out-Null
        } else {
            if ($PSBoundParameters.ContainsKey('Path')) {
                $ConfigFile.Configuration.Add.RemoveAttribute("Path")
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
            Add-Member -InputObject $Result -MemberType NoteProperty -Name "Path" -Value $Path
            Add-Member -InputObject $Result -MemberType NoteProperty -Name "Level" -Value $Level
            $Result
        }
    }
}

Function Get-ODTLogging{
<#
.SYNOPSIS
Gets the value of the Logging section in the configuration file

.PARAMETER TargetFilePath
Required. Full file path for the file.

.Example
Get-ODTLogging -TargetFilePath "$env:Public\Documents\config.xml"
Returns the value of the Logging section if it exists in the specified
file. 

#>
    Param(

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $TargetFilePath

    )

    Process{
        $TargetFilePath = GetFilePath -TargetFilePath $TargetFilePath

        #Load the file
        [System.XML.XMLDocument]$ConfigFile = New-Object System.XML.XMLDocument
        $ConfigFile.Load($TargetFilePath) | Out-Null

        #Check that the file is properly formatted
        if($ConfigFile.Configuration -eq $null){
            throw $NoConfigurationElement
        }
        
        $ConfigFile.Configuration.GetElementsByTagName("Logging") | Select Level, Path
    }

}

Function Remove-ODTLogging{
<#
.SYNOPSIS
Removes the Logging item from configuration xml file

.PARAMETER TargetFilePath
Full file path for the file to be modified and be output to.

.Example
Remove-ODTLogging -TargetFilePath "$env:Public/Documents/config.xml"
Remove the Logging node from the Target File

.Notes
Here is what the portion of configuration file that will be removed by this function:

<Configuration>
  ...
  <Logging Level="Standard" Path="%temp%" />
  ...
</Configuration>

#>
    Param(

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $TargetFilePath

    )

    Process{
        $TargetFilePath = GetFilePath -TargetFilePath $TargetFilePath

        #Load file
        [System.XML.XMLDocument]$ConfigFile = New-Object System.XML.XMLDocument
        $ConfigFile.Load($TargetFilePath) | Out-Null

        $global:saveLastConfigFile = $ConfigFile.OuterXml

        #Check for proper root element
        if($ConfigFile.Configuration -eq $null){
            throw $NoConfigurationElement
        }

        #Get logging element if it exists
        [System.XML.XMLElement]$LoggingElement = $ConfigFile.Configuration.GetElementsByTagName("Logging").Item(0)
        if($ConfigFile.Configuration.Logging -ne $null){
            $ConfigFile.Configuration.removeChild($LoggingElement) | Out-Null
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
            $Result
        }
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

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $Level,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $AcceptEULA,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $TargetFilePath

    )

    Process{
        $TargetFilePath = GetFilePath -TargetFilePath $TargetFilePath

        #Load file
        [System.XML.XMLDocument]$ConfigFile = New-Object System.XML.XMLDocument
        $ConfigFile.Load($TargetFilePath) | Out-Null

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

Function Get-ODTDisplay{
<#
.SYNOPSIS
Gets the value of the Display section in the configuration file

.PARAMETER TargetFilePath
Required. Full file path for the file.

.Example
Get-ODTDisplay -TargetFilePath "$env:Public\Documents\config.xml"
Returns the value of the Display section if it exists in the specified
file. 

#>
    Param(

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $TargetFilePath

    )

    Process{
        $TargetFilePath = GetFilePath -TargetFilePath $TargetFilePath

        #Load the file
        [System.XML.XMLDocument]$ConfigFile = New-Object System.XML.XMLDocument
        $ConfigFile.Load($TargetFilePath) | Out-Null

        #Check that the file is properly formatted
        if($ConfigFile.Configuration -eq $null){
            throw $NoConfigurationElement
        }
        
        $ConfigFile.Configuration.GetElementsByTagName("Display") | Select Level, AcceptEULA
    }

}

Function Remove-ODTDisplay{
<#
.SYNOPSIS
Modifies an existing configuration xml file to remove the diplay item

.PARAMETER TargetFilePath
Full file path for the file to be modified and be output to.

.Example
Remove-ODTDisplay -TargetFilePath "$env:Public/Documents/config.xml"
Sets config show the UI during install

.Notes
Here is what the removed portion of configuration file looks like:

<Configuration>
  ...
  <Display Level="None" AcceptEULA="TRUE" />
  ...
</Configuration>

#>
    Param(

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $TargetFilePath

    )

    Process{
        $TargetFilePath = GetFilePath -TargetFilePath $TargetFilePath

        #Load file
        [System.XML.XMLDocument]$ConfigFile = New-Object System.XML.XMLDocument
        $ConfigFile.Load($TargetFilePath) | Out-Null

        $global:saveLastConfigFile = $ConfigFile.OuterXml

        #Check for proper root element
        if($ConfigFile.Configuration -eq $null){
            throw $NoConfigurationElement
        }

        #Get display element if it exists
        [System.XML.XMLElement]$DisplayElement = $ConfigFile.Configuration.GetElementsByTagName("Display").Item(0)
        if($ConfigFile.Configuration.Display -ne $null){
           $ConfigFile.Configuration.removeChild($LoggingElement) | Out-Null
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

Function LanguagePrompt() {
    Param(
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
           [string] $DefaultLanguage
        )
        
  do {
   Write-Host
   Write-Host "Enter Language (Current: $DefaultLanguage):" -NoNewline
   $selection = Read-Host

   if ($selection) {
     $selection = IsSupportedLanguage -Language $selection
     if (!($selection)) {
       Write-Host "Invalid Selection" -BackgroundColor Red
     } else {
       return $selection
     }
    } else {
      return $DefaultLanguage
    }
  } while($true);
}

Function SelectLanguage() {
  do {
   Write-Host
   Write-Host "Available Language identifiers"
   Write-Host

   $index = 1;
   foreach ($language in $validLanguages) {
      $langSplit = $language.Split("|")

      $lineText = "`t$index - " + $langSplit[0] + " (" + $langSplit[1] + ")"
      Write-Host $lineText
      $index++
   }

   Write-Host
   Write-Host "Select a Language:" -NoNewline
   $selection = Read-Host

   $load = [reflection.assembly]::LoadWithPartialName("'Microsoft.VisualBasic")
   $isNumeric = [Microsoft.VisualBasic.Information]::isnumeric($selection)

   if (!($isNumeric)) {
      Write-Host "Invalid Selection" -BackgroundColor Red
   } else {

     [int] $numSelection = $selection
  
     if ($numSelection -gt 0 -and $numSelection -lt $index) {
        $selectedItem = $validLanguages[$numSelection - 1]
        $langSplit = $selectedItem.Split("|")
        return $langSplit[1]
        break;
     }

     Write-Host "Invalid Selection" -BackgroundColor Red
   }

  } while($true);
}

Function SelectProductId() {
  do {
   Write-Host
   Write-Host "Office Deployment Tool for Click-to-Run Product Ids"
   Write-Host

   $index = 1;
   foreach ($product in $validProductIds) {
      Write-Host "`t$index - $product"
      $index++
   }

   Write-Host
   Write-Host "Select a ProductId:" -NoNewline
   $selection = Read-Host

   $load = [reflection.assembly]::LoadWithPartialName("'Microsoft.VisualBasic")
   $isNumeric = [Microsoft.VisualBasic.Information]::isnumeric($selection)

   if (!($isNumeric)) {
      Write-Host "Invalid Selection" -BackgroundColor Red
   } else {

     [int] $numSelection = $selection

     if ($numSelection -gt 0 -and $numSelection -lt $index) {
        return $validProductIds[$numSelection - 1]
        break;
     }

     Write-Host "Invalid Selection" -BackgroundColor Red
   }

  } while($true);
}

Function SelectBitness() {
  do {
   Write-Host
   Write-Host "Office Bitness"
   Write-Host

   $index = 1;
   Write-Host "`t1 - 32-Bit"
   Write-Host "`t2 - 64-Bit"

   Write-Host
   Write-Host "Select Product Bitness:" -NoNewline
   $selection = Read-Host

   $load = [reflection.assembly]::LoadWithPartialName("'Microsoft.VisualBasic")
   $isNumeric = [Microsoft.VisualBasic.Information]::isnumeric($selection)

   if (!($isNumeric)) {
      Write-Host "Invalid Selection" -BackgroundColor Red
   } else {

     [int] $numSelection = $selection

     if ($numSelection -eq 1 -or $numSelection -eq 2)
     {
        if ($numSelection -eq 1) {
           return "32"
        }
        if ($numSelection -eq 2) {
           return "64"
        }
        break;
     }

     Write-Host "Invalid Selection" -BackgroundColor Red
   }

  } while($true);
}

Function Format-XML ([xml]$xml, $indent=2) { 
    $StringWriter = New-Object System.IO.StringWriter 
    $XmlWriter = New-Object System.XMl.XmlTextWriter $StringWriter 
    $xmlWriter.Formatting = "indented" 
    $xmlWriter.Indentation = $Indent 
    $xml.WriteContentTo($XmlWriter) 
    $XmlWriter.Flush() 
    $StringWriter.Flush() 
    Write-Output $StringWriter.ToString() 
}

Function IsSupportedLanguage() {
    Param(
           [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
           [string] $Language,

           [Parameter()]
           [bool] $ShowLanguages = $true
        )

        $lang = $validLanguages | where {$_.ToString().ToUpper().EndsWith("|$Language".ToUpper())}
          
        if (!($lang)) {
           if ($ShowLanguages) {
              Write-Host
              Write-Host "Invalid or Unsupported Language. Please select a language." -NoNewLine -BackgroundColor Red
              Write-Host

              return SelectLanguage 
           } else {
              throw "Invalid or Unsupported Language: $Language"
           }
           
        }

        return $Language
}

Function IsValidProductId() {
    Param(
           [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
           [string] $ProductId
        )

        $prod = $validProductIds | where {$_.ToString().ToUpper().Equals("$ProductId".ToUpper())}
          
        if (!($prod)) {
            throw "Invalid or Unsupported ProductId: $ProductId"
        }

        return $ProductId
}
