try {
Add-Type  -ErrorAction SilentlyContinue -TypeDefinition @"
   public enum OfficeCTRVersion
   {
      Office2013,
      Office2016
   }
"@
} catch {}

try {
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
         SPDRetail = 16,
         VisioProXVolume = 32,
         VisioStdXVolume = 64,
         ProjectProXVolume = 128,
         ProjectStdXVolume = 256,
     }
}
"
Add-Type -TypeDefinition $enum -ErrorAction SilentlyContinue
} catch {}

try {
$enum2 = "
using System;
 
    [FlagsAttribute]
    public enum LogLevel
    {
        None=0,
        Full=1
    }
"
Add-Type -TypeDefinition $enum2 -ErrorAction SilentlyContinue
} catch {}

function Install-OfficeClickToRun {
    [CmdletBinding()]
    Param(
        [Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true, Position=0)]
        [string] $ConfigurationXML = $NULL,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $TargetFilePath = $NULL,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [OfficeCTRVersion] $OfficeVersion = "Office2016",

        [Parameter()]
        [bool] $WaitForInstallToFinish = $true

    )

    $scriptRoot = GetScriptRoot

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

    [string]$officeCtrPath = ""

    if ($OfficeVersion -eq "Office2013") {
        $officeCtrPath = Join-Path $scriptRoot "Office2013Setup.exe"
        if (!(Test-Path -Path $officeCtrPath)) {
           throw "Cannot find the Office 2013 Setup executable"
        }
    }

    if ($OfficeVersion -eq "Office2016") {
        $officeCtrPath = $scriptRoot + "\Office2016Setup.exe"
        if (!(Test-Path -Path $officeCtrPath)) {
           throw "Cannot find the Office 2016 Setup executable"
        }
    }
    
    if (!($TargetFilePath)) {
      if ($ConfigurationXML) {
         $TargetFilePath = $scriptRoot + "\configuration.xml"
         New-Item -Path $TargetFilePath -ItemType "File" -Value $ConfigurationXML -Force | Out-Null
      }
    }
    
    if (!(Test-Path -Path $TargetFilePath)) {
       $TargetFilePath = $scriptRoot + "\configuration.xml"
    }
    
    $products = Get-ODTProductToAdd -TargetFilePath $TargetFilePath 
    $addNode = Get-ODTAdd -TargetFilePath $TargetFilePath 

    $sourcePath = $addNode.SourcePath
    $version = $addNode.Version
    $edition = $addNode.OfficeClientEdition

    foreach ($product in $products)
    {
        if ($product) {
          $languages = getProductLanguages -Product $product 
          $existingLangs = checkForLanguagesInSourceFiles -Languages $languages -SourcePath $sourcePath -Version $version -Edition $edition
          if ($product.ProductId) {
              Set-ODTProductToAdd -TargetFilePath $TargetFilePath -ProductId $product.ProductId -LanguageIds $existingLangs | Out-Null
          }
        }
    }

    $cmdLine = $officeCtrPath
    $cmdArgs = "/configure " + $TargetFilePath

    Write-Host "Installing Office Click-To-Run..."

    StartProcess -execFilePath $cmdLine -execParams $cmdArgs -WaitForExit $false

    if ($WaitForInstallToFinish) {
         Wait-ForOfficeCTRInstall
    }
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

    $scriptRoot = GetScriptRoot

    $returnLanguages = @()

    if (!($SourcePath)) {
      $localSource = $scriptRoot + "\Office\Data"
      if (Test-Path -Path $localSource) {
         $SourcePath = $scriptRoot
      }
    }

    if (!($Version)) {
       $localPath = $env:TEMP
       $cabPath = $scriptRoot + "\Office\Data\v$Edition.cab"
       $cabFolderPath = $scriptRoot + "\Office\Data"
       $vdXmlPath = $localPath + "\VersionDescriptor.xml"
       
       if (Test-Path -Path $cabPath) {
          Invoke-Expression -Command "Expand $cabPath -F:VersionDescriptor.xml $localPath" | Out-Null
          $Version = getVersionFromVersionDescriptor -vesionDescriptorPath $vdXmlPath
          Remove-Item -Path $vdXmlPath -Force
       }
    }

    $verionDir = $scriptRoot + "\Office\Data\$Version"
    
    if (Test-Path -Path $verionDir) {
       foreach ($lang in $Languages) {
          $fileName = "stream.x86.$lang.dat"
          if ($Edition -eq "64") {
             $fileName = "stream.x64.$lang.dat"
          }
          
          $langFile = $verionDir + "\" + $fileName 
          
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
      if (!($languages -contains ($language))) {
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
          if (!($languages -contains $language)) {
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
           $content = Get-Content $TargetFilePath
           $ConfigFile.LoadXml($content) | Out-Null
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
                $Result = New-Object 傍ypeName PSObject 

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
            if ($ProductId) {
            

                [System.XML.XMLElement]$ProductElement = $ConfigFile.Configuration.Add.Product | where { $_.ID -eq $ProductId }
                if ($ProductElement) {
                $tempId = $ProductElement.GetAttribute("ID")
                
                
                $Result = New-Object 傍ypeName PSObject 
                Add-Member -InputObject $Result -MemberType NoteProperty -Name "ProductId" -Value $tempId 
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
           $content = Get-Content $TargetFilePath
           $ConfigFile.LoadXml($content) | Out-Null
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
        [bool] $AcceptEULA = $true,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $TargetFilePath

    )

    Process{
        $TargetFilePath = GetFilePath -TargetFilePath $TargetFilePath

        #Load file
        [System.XML.XMLDocument]$ConfigFile = New-Object System.XML.XMLDocument

        if ($TargetFilePath) {
           $content = Get-Content $TargetFilePath
           $ConfigFile.LoadXml($content) | Out-Null
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
        if($Level -ne $null){
            $DisplayElement.SetAttribute("Level", $Level) | Out-Null
        } else {
            if ($PSBoundParameters.ContainsKey('Level')) {
                $ConfigFile.Configuration.Add.RemoveAttribute("Level")
            }
        }

        if($AcceptEULA -ne $null){
            $DisplayElement.SetAttribute("AcceptEULA", $AcceptEULA.ToString().ToUpper()) | Out-Null
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
            $Result = New-Object 傍ypeName PSObject 
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
    
   $locationPath = (Get-Location).Path
    
    if (!($TargetFilePath.IndexOf('\') -gt -1)) {
      $TargetFilePath = $locationPath + "\" + $TargetFilePath
    }

    return $TargetFilePath
}

Function Get-OfficeCTRRegPath() {
    $path15 = 'SOFTWARE\Microsoft\Office\15.0\ClickToRun'
    $path16 = 'SOFTWARE\Microsoft\Office\ClickToRun'

    if (Test-Path "HKLM:\$path15") {
      return $path15
    } else {
      if (Test-Path "HKLM:\$path16") {
         return $path16
      }
    }
}

Function Set-ODTProductToAdd{
<#
.SYNOPSIS
Modifies an existing configuration xml file to modify a existing product item.

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
        [Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true, Position=0)]
        [string] $ConfigurationXML = $NULL,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $TargetFilePath = $NULL,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [Microsoft.Office.Products] $ProductId = "Unknown",

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [Alias("LanguageId")]
        [string[]] $LanguageIds = $NULL,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string[]] $ExcludeApps = $NULL

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
        }

        #Load the file
        [System.XML.XMLDocument]$ConfigFile = New-Object System.XML.XMLDocument
        
        if ($TargetFilePath) {
           $content = Get-Content $TargetFilePath
           $ConfigFile.LoadXml($content) | Out-Null
        } else {
            if ($ConfigurationXml) 
            {
              $ConfigFile.LoadXml($ConfigurationXml) | Out-Null
              $global:saveLastConfigFile = $NULL
              $global:saveLastFilePath = $NULL
              $TargetFilePath = $NULL
            }
        }

        $global:saveLastConfigFile = $ConfigFile.OuterXml

        #Check that the file is properly formatted
        if($ConfigFile.Configuration -eq $null){
            throw $NoConfigurationElement
        }

        [System.XML.XMLElement]$AddElement=$NULL
        if($ConfigFile.Configuration.Add -eq $null){
           throw "Cannot find 'Add' element"
        }

        $AddElement = $ConfigFile.Configuration.Add 

        #Set the desired values
        [System.XML.XMLElement]$ProductElement = $ConfigFile.Configuration.Add.Product | ?  ID -eq $ProductId
        if($ProductElement -eq $null){
           throw "Cannot find Product with Id '$ProductId'"
        }

        if ($LanguageIds) {
            $existingLangs = $ProductElement.selectnodes("./Language")
            if ($existingLangs.count -gt 0) {
                foreach ($lang in $existingLangs) {
                  $ProductElement.removeChild($lang) | Out-Null
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
        }

        if ($ExcludeApps) {
            $existingExcludes = $ProductElement.selectnodes("./ExcludeApp")
            if ($existingExcludes.count -gt 0) {
                foreach ($exclude in $existingLangs) {
                  $ProductElement.removeChild($exclude) | Out-Null
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
            $Result = New-Object 傍ypeName PSObject 
            Add-Member -InputObject $Result -MemberType NoteProperty -Name "TargetFilePath" -Value $TargetFilePath
            Add-Member -InputObject $Result -MemberType NoteProperty -Name "ProductId" -Value $ProductId
            Add-Member -InputObject $Result -MemberType NoteProperty -Name "LanguageIds" -Value $LanguageIds
            $Result
        }


    }

}

Function Wait-ForOfficeCTRInstall() {
    [CmdletBinding()]
    Param(
        [Parameter()]
        [int] $TimeOutInMinutes = 120
    )

    begin {
        $HKLM = [UInt32] "0x80000002"
        $HKCR = [UInt32] "0x80000000"
    }

    process {
       Write-Host "Waiting for Update to Complete..."

       Start-Sleep -Seconds 20

       $mainRegPath = Get-OfficeCTRRegPath 
       $scenarioPath = $mainRegPath + "\scenario"

       $regProv = Get-Wmiobject -list "StdRegProv" -namespace root\default -ErrorAction Stop

       [DateTime]$startTime = Get-Date

       [string]$executingScenario = ""
       $failure = $false
       $updateRunning=$false
       [string[]]$trackProgress = @()
       [string[]]$trackComplete = @()
       do {
           $allComplete = $true
           try {
              $exScenario = $regProv.GetStringValue($HKLM, $mainRegPath, "ExecutingScenario")
           } catch { }
           if ($exScenario) {
              $executingScenario = $exScenario.sValue
           }

           $scenarioKeys = $regProv.EnumKey($HKLM, $scenarioPath)
           foreach ($scenarioKey in $scenarioKeys.sNames) {
              if (!($executingScenario)) { continue }
              if ($scenarioKey.ToLower() -eq $executingScenario.ToLower()) {
                $taskKeyPath = $scenarioPath + "\$scenarioKey\TasksState"
                $taskValues = $regProv.EnumValues($HKLM, $taskKeyPath).sNames

                foreach ($taskValue in $taskValues) {
                    [string]$status = $regProv.GetStringValue($HKLM, $taskKeyPath, $taskValue).sValue
                    $operation = $taskValue.Split(':')[0]
                    $keyValue = $taskValue

                    if ($status.ToUpper() -eq "TASKSTATE_FAILED") {
                        $failure = $true
                    }

                    $displayValue = showTaskStatus -Operation $operation -Status $status -DateTime (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')

                    if (($status.ToUpper() -eq "TASKSTATE_COMPLETED") -or`
                        ($status.ToUpper() -eq "TASKSTATE_CANCELLED") -or`
                        ($status.ToUpper() -eq "TASKSTATE_FAILED")) {
                        if (($trackProgress -contains $keyValue) -and !($trackComplete -contains $keyValue) ) {
                            $displayValue
                            $trackComplete += $keyValue 
                        }
                    } else {
                        $allComplete = $false
                        $updateRunning=$true

                        if (!$trackProgress -contains $keyValue) {
                            $trackProgress += $keyValue 
                            $displayValue
                        }
                    }
                }
              }
           }


           if ($allComplete) {
              break;
           }

           if ($startTime -lt (Get-Date).AddHours(-$TimeOutInMinutes)) {
              throw "Waiting for Update Timed-Out"
              break;
           }

           Start-Sleep -Seconds 5
       } while($true -eq $true) 

       if ($updateRunning) {
          if ($failure) {
            Write-Host "Update Failed"
          } else {
            Write-Host "Update Complete"
          }
       } else {
          Write-Host "Update Not Running"
       } 
    }
}

function showTaskStatus() {
    [CmdletBinding()]
    Param(
        [Parameter()]
        [string] $Operation = "",

        [Parameter()]
        [string] $Status = "",

        [Parameter()]
        [string] $DateTime = ""
    )

    $results = new-object PSObject[] 0;
    $Result = New-Object 傍ypeName PSObject 
    Add-Member -InputObject $Result -MemberType NoteProperty -Name "Operation" -Value $Operation
    Add-Member -InputObject $Result -MemberType NoteProperty -Name "Status" -Value $Status
    Add-Member -InputObject $Result -MemberType NoteProperty -Name "DateTime" -Value $DateTime
    return $Result
}

Function StartProcess {
	Param
	(
        [Parameter()]
		[String]$execFilePath,

        [Parameter()]
        [String]$execParams,

        [Parameter()]
        [bool]$WaitForExit = $false
	)

    Try
    {
        $startExe = new-object System.Diagnostics.ProcessStartInfo
        $startExe.FileName = $execFilePath
        $startExe.Arguments = $execParams
        $startExe.CreateNoWindow = $false
        $startExe.UseShellExecute = $false

        $execStatement = [System.Diagnostics.Process]::Start($startExe) 
        if ($WaitForExit) {
           $execStatement.WaitForExit()
        }
    }
    Catch
    {
        Write-Log -Message $_.Exception.Message -severity 1 -component "Office 365 Update Anywhere"
    }
}

Function GetScriptRoot() {
 process {
     [string]$scriptPath = "."

     if ($PSScriptRoot) {
       $scriptPath = $PSScriptRoot
     } else {
       $scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
       $scriptPath = (Get-Item -Path ".\").FullName
     }
     return $scriptPath
 }
}

Function GetScriptPath() {
 return GetScriptRoot
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
