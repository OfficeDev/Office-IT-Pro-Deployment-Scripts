function Generate-ODTLanguagePackXML {
<#
.SYNOPSIS
Create an ODT configuration file to deploy additional language packs

.DESCRIPTION
This script will create a new xml file that should be used to deploy additional 
language packs to computers with Office 365 ProPlus already installed.

.PARAMETER TargetFilePath
The full path where to save the file.

.PARAMETER OfficeClientEdition
The bit of Office. Choose between 32 and 64.

.PARAMETER Languages
The list of available languages.

.EXAMPLE
Generate-ODTLanguagePackXML -TargetFilePath $env:temp\LanguagePacks.xml -Languages de-de,es-es,fr-fr -OfficeClientEdition 64
A new xml file will be created in the temp directory called LanguagePacks.xml which will be used to install the 64-bit
editions of German, Spanish, and French language packs.

.EXAMPLE
Generate-ODTLanguagePackXML -TargetFilePath $env:temp\LanguagePacks.xml -Languages de-de,es-es,fr-fr | fl
A new xml file will be created in the temp directory called LanguagePacks.xml which will be used to install the 32-bit
editions of German, Spanish, and French language packs. The output of the xml file will be displayed on the PowerShell console.

.NOTES
Date created: 03-02-2017
Date modified: 03-02-2017
#>
[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [Parameter(ValueFromPipelineByPropertyName=$true)]
    [String]$TargetFilePath = $NULL,

    [Parameter(ValueFromPipelineByPropertyName=$true)]
    [ValidateSet("32","64")]
    [string]$OfficeClientEdition = '32',

    [Parameter(ValueFromPipelineByPropertyName=$true)]
    [ValidateSet("en-us","MatchOS","ar-sa","bg-bg","zh-cn","zh-tw","hr-hr","cs-cz","da-dk","nl-nl","et-ee","fi-fi","fr-fr","de-de","el-gr","he-il","hi-in","hu-hu","id-id","it-it",
                "ja-jp","kk-kz","ko-kr","lv-lv","lt-lt","ms-my","nb-no","pl-pl","pt-br","pt-pt","ro-ro","ru-ru","sr-latn-rs","sk-sk","sl-si","es-es","sv-se","th-th",
                "tr-tr","uk-ua","vi-vn")] 
    [string[]]$Languages
)

begin {   
    [string]$tempStr = $MyInvocation.MyCommand.Path
    $scriptPath = GetScriptPath     
}

process{
    if ($TargetFilePath) {
        $folderPath = Split-Path -Path $TargetFilePath -Parent
        $fileName = Split-Path -Path $TargetFilePath -Leaf
        if ($folderPath) {
            [system.io.directory]::CreateDirectory($folderPath) | Out-Null
        }
    }

    [System.XML.XMLDocument]$ConfigFile = New-Object System.XML.XMLDocument
   
    #Generate the language pack xml file
    odtAddLanguagePackProduct -ConfigDoc $ConfigFile -Platform $OfficeClientEdition -LanguageIds $Languages

    $formattedXml = Format-XML ([xml]($ConfigFile)) -indent 4

    if ($TargetFilePath) {
       $formattedXml | Out-File -FilePath $TargetFilePath
    }

    #Display the results to the console
    $Result = New-Object -TypeName PSObject 
    Add-Member -InputObject $Result -MemberType NoteProperty -Name "TargetFilePath" -Value $TargetFilePath
    Add-Member -InputObject $Result -MemberType NoteProperty -Name "LanguageIds" -Value $Languages
    Add-Member -InputObject $Result -MemberType NoteProperty -Name "ConfigurationXML" -Value $formattedXml
    $Result
}

}

function odtAddLanguagePackProduct() {
    param(
       [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
       [System.XML.XMLDocument]$ConfigDoc = $NULL,

       [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
       [string]$ProductId = "LanguagePack",

       [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
       [string]$Platform = $NULL,

       [Parameter(ValueFromPipelineByPropertyName=$true)]
       [string[]]$LanguageIds = @()
    )

    [System.XML.XMLElement]$ConfigElement=$NULL
    if($ConfigDoc.Configuration -eq $null){
        $ConfigElement=$ConfigDoc.CreateElement("Configuration")
        $ConfigDoc.appendChild($ConfigElement) | Out-Null
    }

    [System.XML.XMLElement]$AddElement=$NULL
    if($ConfigFile.Configuration.Add -eq $null){
        $AddElement=$ConfigDoc.CreateElement("Add")
        $ConfigDoc.DocumentElement.appendChild($AddElement) | Out-Null
    } else {
        $AddElement = $ConfigDoc.Configuration.Add 
    }

    if ($Platform) {
       $AddElement.SetAttribute("OfficeClientEdition", $Platform) | Out-Null
    }

    [System.XML.XMLElement]$ProductElement = $ConfigDoc.Configuration.Add.Product | where { $_.ID -eq $ProductId }
    if($ProductId){
        if($ProductElement -eq $null){
            [System.XML.XMLElement]$ProductElement=$ConfigDoc.CreateElement("Product")
            $AddElement.appendChild($ProductElement) | Out-Null
            $ProductElement.SetAttribute("ID", $ProductId) | Out-Null
        }
    }

    foreach($LanguageId in $LanguageIds){    
       if ($LanguageId) {
          if ($LanguageId.Length -gt 0) {
            [System.XML.XMLElement]$LanguageElement = $ProductElement.Language | where { $_.ID -eq $LanguageId }
            if($LanguageElement -eq $null){
                [System.XML.XMLElement]$LanguageElement=$ConfigFile.CreateElement("Language")
                $ProductElement.appendChild($LanguageElement) | Out-Null
                $LanguageElement.SetAttribute("ID", $LanguageId.ToString().ToLower()) | Out-Null
            }
          }
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
        [Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true, Position=0)]
        [string] $ConfigurationXML = $NULL,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string[]] $ProductId = "Unknown",

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $TargetFilePath,

        [Parameter(ParameterSetName="All", ValueFromPipelineByPropertyName=$true)]
        [switch] $All
    )

    Process{
        $TargetFilePath = GetFilePath -TargetFilePath $TargetFilePath

        foreach($Product in $ProductId){
            if ($Product -eq "Unknown") {
                $Product = SelectProductId
            }

            $Product = IsValidProductId -ProductId $Product

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

            $global:saveLastConfigFile = $ConfigFile.OuterXml

            #Check that the file is properly formatted
            if($ConfigFile.Configuration -eq $null){
                throw $NoConfigurationElement
            }

            if($ConfigFile.Configuration.Add -eq $null){
                throw $NoAddElement
            }

            if (!($All)) {
                #Set the desired values
                [System.XML.XMLElement]$ProductElement = $ConfigFile.Configuration.Add.Product | Where { $_.ID -eq $Product }
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
               $ConfigFile.Configuration.Add.RemoveAll() | Out-Null
           
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
}

Function GetScriptPath() {
    [string]$scriptPath = "."
    
    if ($PSScriptRoot) {
      $scriptPath = $PSScriptRoot
    } else {
      $scriptPath = (Get-Item -Path ".\").FullName
    }
    
    return $scriptPath
}

function Format-XML ([xml]$xml, $indent=2) { 
    $StringWriter = New-Object System.IO.StringWriter 
    $XmlWriter = New-Object System.XMl.XmlTextWriter $StringWriter 
    $xmlWriter.Formatting = "indented" 
    $xmlWriter.Indentation = $Indent 
    $xml.WriteContentTo($XmlWriter) 
    $XmlWriter.Flush() 
    $StringWriter.Flush() 
    Write-Output $StringWriter.ToString() 
}