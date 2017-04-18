[CmdletBinding(SupportsShouldProcess=$true)]
param(
[Parameter(ValueFromPipelineByPropertyName=$true)]
[bool]$RemoveClickToRunVersions = $false,

[Parameter(ValueFromPipelineByPropertyName=$true)]
[bool]$Remove2016Installs = $false,

[Parameter(ValueFromPipelineByPropertyName=$true)]
[bool]$Force = $false,

[Parameter(ValueFromPipelineByPropertyName=$true)]
[bool]$KeepUserSettings = $true,

[Parameter(ValueFromPipelineByPropertyName=$true)]
[bool]$KeepLync = $false,

[Parameter(ValueFromPipelineByPropertyName=$true)]
[bool]$NoReboot = $false,

[Parameter(ValueFromPipelineByPropertyName=$true)]
[bool]$Quiet = $true,

[Parameter()]
[ValidateSet("AllOfficeProducts","MainOfficeProduct","Visio","Project")]
[string[]]$ProductsToRemove = "AllOfficeProducts",

[Parameter()]
[ValidateSet("O365ProPlusRetail","O365BusinessRetail","VisioProRetail","ProjectProRetail", "SPDRetail", "VisioProXVolume", "VisioStdXVolume", 
             "ProjectProXVolume", "ProjectStdXVolume", "InfoPathRetail", "SkypeforBusinessEntryRetail", "LyncEntryRetail")]
[string]$C2RProductsToRemove = "O365ProPlusRetail",

[Parameter()]
[string]$LogFilePath
)

$validProductIds = @("O365ProPlusRetail","O365BusinessRetail","VisioProRetail","ProjectProRetail", "SPDRetail", "VisioProXVolume", "VisioStdXVolume", 
                     "ProjectProXVolume", "ProjectStdXVolume", "InfoPathRetail", "SkypeforBusinessEntryRetail", "LyncEntryRetail")

$validLanguages = @(
"English|en-us",          #beginning of core languages
"MatchOS|MatchOS",
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
"Kazakh|kk-kz",
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
"Ukrainian|uk-ua",
"Vietnamese|vi-vn",       #end of core languages
"Afrikaans (South Africa)|af-za",                #beginning of partial languages
"Albanian (Albania)|sq-al",
"Amharic (Ethiopia)|am-et",
"Armenian (Armenia)|hy-am",
"Assamese (India)|as-in",
"Azerbaijani (Latin, Azerbaijan)|az-latn-az",
"Basque (Basque)|eu-es",
"Belarusian (Belarus)|be-by",
"Bangla (Bangladesh)|bn-bd",
"Bangla (India)|bn-in",
"Bosnian (Latin, Bosnia and Herzegovina)|bs-latn-ba",
"Catalan (Catalan)|ca-es",
"Dari (Afghanistan)|prs-af",
"Filipino (Philippines)|fil-ph",
"Galician (Galician)|gl-es",
"Georgian (Georgia)|ka-ge",
"Gujarati (India)|gu-in",
"Icelandic (Iceland)|is-is",
"Irish (Ireland)|ga-ie",
"Kannada (India)|kn-in",
"Khmer (Cambodia)|km-kh",
"Kiswahili (Kenya)|sw-ke",
"Konkani (India)|kok-in",
"Kyrgyz (Kyrgyzstan)|ky-kg",
"Luxembourgish (Luxembourg)|lb-lu",
"Macedonian (Former Yugoslav Republic of Macedonia)|mk-mk",
"Malayalam (India)|ml-in",
"Maltese (Malta)|mt-mt",
"Maori (New Zealand)|mi-nz",
"Marathi (India)|mr-in",
"Mongolian (Cyrillic, Mongolia)|mn-mn",
"Nepali (Nepal)|ne-np",
"Norwegian, Nynorsk (Norway)|nn-no",
"Odia (India)|or-in",
"Persian (Iran)|fa-ir",
"Punjabi (India)|pa-in",
"Quechua (Peru)|quz-pe",
"Scottish Gaelic (United Kingdom)|gd-gb",
"Serbian (Cyrillic, Serbia)|sr-cyrl-rs",
"Serbian (Cyrillic, Bosnia and Herzegovina)|sr-cyrl-ba",
"Sindhi (Islamic Republic of Pakistan)|sd-arab-pk",
"Sinhala (Sri Lanka)|si-lk",
"Tamil (India)|ta-in",
"Tatar (Russia)|tt-ru",
"Telugu (India)|te-in",
"Turkmen (Turkmenistan)|tk-tm",
"Urdu (Islamic Republic of Pakistan)|ur-pk",
"Uyghur (PRC)|ug-cn",
"Uzbek (Latin, Uzbekistan)|uz-latn-uz",
"Valencian (Spain)|ca-es-valencia",
"Welsh (United Kingdom)|cy-gb",         #end of partial languages
"Hausa (Latin, Nigeria)|ha-latn-ng",    #beginning of proofing languages
"Igbo (Nigeria)|ig-ng",
"isiXhosa (South Africa)|xh-za",
"isiZulu (South Africa)|zu-za",
"Kinyarwanda (Rwanda)|rw-rw",
"Pashto (Afghanistan)|ps-af",
"Romansh (Switzerland)|rm-ch",
"Sesotho sa Leboa (South Africa)|nso-za",
"Setswana (South Africa)|tn-za",
"Wolof (Senegal)|wo-sn",
"Yoruba (Nigeria)|yo-ng")

Function Remove-PreviousOfficeInstalls{
<#
.SYNOPSIS
Automate the process to remove Office products.

.DESCRIPTION
Automate the process to remove Office products.

.PARAMETER RemoveClickToRunVersions
Set the value to $true to also remove Click-To-Run version of Office.

.PARAMETER Remove2016Installs
Set the value to $true to also remove 2016 versions of Office.

.PARAMETER Force
Set the value to $true to force an uninstall.

.PARAMETER KeepUserSettings
By default, the value is set to $true. Set to $false to remove user settings.

.PARAMETER KeepLync
Set the value to $true to preserve the Lync installation.

.PARAMETER NoReboot
By default, the value is set to $false. Set to $true to offer the reboot prompt if needed.
 
.PARAMETER Quiet
By default, the value is set to $true. Set to $false to show the progress of 
the uninstall.

.PARAMETER ProductsToRemove
By default the value is AllOfficeProducts which will remove all Office products. Set this value
to MainOfficeProduct, Visio, and/or Project to only remove the specified product.

.EXAMPLE
Remove-PreviousOfficeInstalls
In this example all Office products, except for click to run or 2016, will be removed.

.EXAMPLE
Remove-PreviousOfficeInstalls -ProductsToRemove MainOfficeProduct,Visio
In this example the primary office product and Visio will be removed.Click-To-Run or 2016
products will not be removed.

.EXAMPLE
Remove-PreviousOfficeInstalls -ProductsToRemove MainOfficeProduct -RemoveClickToRunVersions $true
In this example the primary Office product will be removed even if it is Click-To-Run.

#>
  [CmdletBinding(SupportsShouldProcess=$true)]
  param(
    [Parameter(ValueFromPipelineByPropertyName=$true)]
    [bool]$RemoveClickToRunVersions = $true,

    [Parameter(ValueFromPipelineByPropertyName=$true)]
    [bool]$Remove2016Installs = $false,

    [Parameter(ValueFromPipelineByPropertyName=$true)]
    [bool]$Force = $false,

    [Parameter(ValueFromPipelineByPropertyName=$true)]
    [bool]$KeepUserSettings = $true,

    [Parameter(ValueFromPipelineByPropertyName=$true)]
    [bool]$KeepLync = $false,

    [Parameter(ValueFromPipelineByPropertyName=$true)]
    [bool]$NoReboot = $false,

    [Parameter(ValueFromPipelineByPropertyName=$true)]
    [bool]$Quiet = $true,

    [Parameter(ValueFromPipelineByPropertyName=$true)]
    [ValidateSet("AllOfficeProducts","MainOfficeProduct","Visio","Project")]
    [string[]]$ProductsToRemove = "AllOfficeProducts",

    [Parameter()]
    [string]$LogFilePath
  )

  Process {
    $currentFileName = Get-CurrentFileName
    Set-Alias -name LINENUM -value Get-CurrentLineNumber

    $c2rVBS = "OffScrubc2r.vbs"
    $03VBS = "OffScrub03.vbs"
    $07VBS = "OffScrub07.vbs"
    $10VBS = "OffScrub10.vbs"
    $15MSIVBS = "OffScrub_O15msi.vbs"
    $16MSIVBS = "OffScrub_O16msi.vbs"

    $argList = ""
    $MainArgListProducts = @()
    $VisioArgListProducts = @()
    $ProjectArgListProducts = @()

    $officeProducts = Get-OfficeVersion -ShowAllInstalledProducts | select *

    [bool]$isVisioC2R = $false
    [bool]$isProjectC2R = $false
   
    if($ProductsToRemove -eq 'AllOfficeProducts'){
        $argListProducts += "CLIENTALL"
    } else {       
        foreach($product in $ProductsToRemove){
            switch($product){
                "MainOfficeProduct"{
                    $OfficeProduct = GetProductName -ProductName MainOfficeProduct
                    WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "OfficeProduct set to $OfficeProduct" -LogFilePath $LogFilePath
                    $MainOfficeProduct = $OfficeProduct | ? {$_.DisplayName -notmatch "Language Pack"}
                    $OfficeLanguagePacks = $officeProduct | ? {$_.DisplayName -match "Language Pack"}
                    if($OfficeLanguagePacks){
                        foreach($OffLang in $OfficeLanguagePacks){
                            $OfficeLanguagePacks += $OffLang.Name
                        }
                    }
                    $OfficeArgListProducts += $MainOfficeProduct.Name
                    $OfficeArgListProducts = $OfficeArgListProducts -join ","
                }
                "Visio" {
                    $VisioProduct = GetProductName -ProductName Visio
                    WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "VisioProduct set to $VisioProduct" -LogFilePath $LogFilePath
                    $MainVisioProduct = $VisioProduct | ? {$_.DisplayName -notmatch "Language Pack"}
                    $VisioLanguagePacks = $VisioProduct | ? {$_.DisplayName -match "Language Pack"}
                    if($VisioLanguagePacks){
                        foreach($VisLang in $VisioLanguagePacks){
                            $VisioArgListProducts += $VisLang.Name
                        }
                    }
                    $VisioArgListProducts += $MainVisioProduct.Name
                    $VisioArgListProducts = $VisioArgListProducts -join ","

                    foreach($product in $officeProducts){
                        if($product.DisplayName.ToLower() -eq $VisioProduct.DisplayName.ToLower()){
                            $VisioProdName = $product
                        }
                    }

                    if($VisioProdName.ClickToRun -eq $true){
                        $isVisioC2R = $true
                    }
                }
                "Project" {
                    $ProjectProduct = GetProductName -ProductName Project
                    WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "ProjectProduct set to $ProjectProduct" -LogFilePath $LogFilePath
                    $MainProjectProduct = $ProjectProduct | ? {$_.DisplayName -notmatch "Language Pack"}
                    $ProjectLanguagePacks = $ProjectProduct | ? {$_.DisplayName -match "Language Pack"}
                    if($ProjectLanguagePacks){
                        foreach($ProjLang in $ProjectLanguagePacks){
                            $ProjectArgListProducts += $ProjLang.Name
                        }
                    }
                    $ProjectArgListProducts += $MainProjectProduct.Name
                    $ProjectArgListProducts = $ProjectArgListProducts -join ","

                    foreach($product in $officeProducts){
                        if($product.DisplayName.ToLower() -eq $ProjectProduct.DisplayName.ToLower()){
                            $ProjectProdName = $product
                        }
                    }

                    if($ProjectProdName.ClickToRun -eq $true){
                        $isProjectC2R = $true
                    }
                }
            }
        }
    }

    if($Quiet){
        $argList += " /QUIET"
    }
       
    if ($Force) {
        $argList += " /FORCE"
    }

    if ($KeepUserSettings) {
       $argList += " /KEEPUSERSETTINGS"
    } else {
       $argList += " /DELETEUSERSETTINGS"
    }

    if ($KeepLync) {
       $argList += " /KEEPLYNC"
    } else {
       $argList += " /REMOVELYNC"
    }

    if ($NoReboot) {
        $argList += " /NOREBOOT"
    }

    $scriptPath = GetScriptRoot

    Write-Host "Detecting Office installs..."
    WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Detecting Office installs..." -LogFilePath $LogFilePath

    $officeVersions = Get-OfficeVersion -ShowAllInstalledProducts | select *
    $ActionFiles = @()
    
    $removeOffice = $true
    if (!( $officeVersions)) {
       Write-Host "Microsoft Office is not installed"
       WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Microsoft Office is not installed" -LogFilePath $LogFilePath
       $removeOffice = $false
    }

    if ($removeOffice) {
        [bool]$office03Removed = $false
        [bool]$office07Removed = $false
        [bool]$office10Removed = $false
        [bool]$office15Removed = $false
        [bool]$office16Removed = $false
        [bool]$officeC2RRemoved = $false

        [bool]$c2r2013Installed = $false
        [bool]$c2r2016Installed = $false
                
        if($ProductsToRemove -ne 'AllOfficeProducts'){
            foreach($product in $ProductsToRemove){
                switch($product){
                    "MainOfficeProduct" {
                        $MainOfficeProductName = $MainOfficeProduct.Name
                        
                        if((Get-OfficeVersion | select *).ClickToRun -eq $true){
                            $c2rInstalled = $true
                        }

                        switch($MainOfficeProduct.Version){
                            "11" {
                                $ActionFile = "$scriptPath\$03VBS"
                            }
                            "12" {
                                $ActionFile = "$scriptPath\$07VBS"
                            }
                            "14" {
                                $ActionFile = "$scriptPath\$10VBS"
                            }
                            "15" {             
                                if(!$c2rInstalled){
                                    $ActionFile = "$scriptPath\$15MSIVBS"
                                } else {
                                    if($RemoveClickToRunVersions){
                                        Remove-OfficeClickToRun -C2RProductsToRemove $MainOfficeProduct.Name.Split('-')[0].Trim()
                                    } else {
                                        WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Office 2013 Click-To-Run cannot be removed. Use the -RemoveClickToRunVersions parameter to remove Click-To-Run 2016 installs." -LogFilePath $LogFilePath
                                        throw "Office 2013 Click-To-Run cannot be removed. Use the -RemoveClickToRunVersions parameter to remove Click-To-Run 2016 installs."
                                    }
                                }
                            }
                            "16" {
                                if($Remove2016Installs){
                                    if(!$c2rInstalled){
                                        $ActionFile = "$scriptPath\$16MSIVBS"
                                    } else {
                                        if($RemoveClickToRunVersions){
                                            Remove-OfficeClickToRun -C2RProductsToRemove $MainOfficeProduct.Name.Split('-')[0].Trim()
                                        } else {
                                            WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Office 2016 Click-To-Run cannot be removed. Use the -RemoveClickToRunVersions parameter to remove Click-To-Run 2016 installs." -LogFilePath $LogFilePath
                                            throw "Office 2016 Click-To-Run cannot be removed. Use the -RemoveClickToRunVersions parameter to remove Click-To-Run 2016 installs."
                                        }
                                    }
                                } else {
                                    WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Office 2016 Click-To-Run cannot be removed. Use the -RemoveClickToRunVersions and -Remove2016Installs parameters to remove Click-To-Run 2016 installs." -LogFilePath $LogFilePath
                                    throw "Office 2016 Click-To-Run cannot be removed. Use the -RemoveClickToRunVersions and -Remove2016Installs parameters to remove Click-To-Run 2016 installs."
                                }
                            }
                        }

                        try{
                             if($ActionFile -And (Test-Path -Path $ActionFile)){
                                $MainOfficeProductDisplayName = $MainOfficeProduct.DisplayName
                                Write-Host "`tRemoving "$MainOfficeProduct.DisplayName"..."
                                WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Removing the MainOfficeProduct..." -LogFilePath $LogFilePath
                                $cmdLine = """$ActionFile"" $MainOfficeProductName $argList"
                                $cmd = "cmd /c cscript //Nologo $cmdLine"
                                Invoke-Expression $cmd
                            } else {
                                throw "Required file missing: $ActionFile"
                            }
                        } catch {}                                  
                    }
                    "Visio" {
                        Write-Host "`tRemoving Visio products..."
                        WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Removing Visio products..." -LogFilePath $LogFilePath
                        $VisioProductName = $VisioProduct.Name

                        switch($VisioProduct.Version){
                            "11" {
                                $ActionFile = "$scriptPath\$03VBS"
                            }
                            "12" {
                                $ActionFile = "$scriptPath\$07VBS"
                            }
                            "14" {
                                $ActionFile = "$scriptPath\$10VBS"
                            }
                            "15" {
                                if(!$isVisioC2R){
                                    $ActionFile = "$scriptPath\$15MSIVBS"
                                } else {
                                    if($RemoveClickToRunVersions){
                                        Remove-OfficeClickToRun -C2RProductsToRemove "VisioProRetail","VisioProXVolume", "VisioStdXVolume"
                                    } else {
                                        WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Visio cannot be removed. Use the -RemoveClickToRunVersions parameter to remove Click-To-Run 2016 installs." -LogFilePath $LogFilePath
                                        throw "Visio cannot be removed. Use the -RemoveClickToRunVersions parameter to remove Click-To-Run 2016 installs."
                                    }
                                }
                            }
                            "16" {
                                if($Remove2016Installs){
                                    if(!$isVisioC2R){
                                        $ActionFile = "$scriptPath\$16MSIVBS"
                                    } else {
                                        if($RemoveClickToRunVersions){
                                            Remove-OfficeClickToRun -C2RProductsToRemove "VisioProRetail","VisioProXVolume", "VisioStdXVolume"
                                        } else {
                                            WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Visio cannot be removed. Use the -RemoveClickToRunVersions parameter to remove Click-To-Run 2016 installs." -LogFilePath $LogFilePath
                                            throw "Visio cannot be removed. Use the -RemoveClickToRunVersions parameter to remove Click-To-Run 2016 installs."
                                        }
                                    }
                                } else {
                                    WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Visio cannot be removed. Use the -RemoveClickToRunVersions and -Remove2016Installs parameters to remove Click-To-Run 2016 installs." -LogFilePath $LogFilePath
                                    throw "Visio cannot be removed. Use the -RemoveClickToRunVersions and -Remove2016Installs parameters to remove Click-To-Run 2016 installs."
                                }
                            }
                        }

                        if($ActionFile -And (Test-Path -Path $ActionFile)){
                            $cmdLine = """$ActionFile"" $VisioArgListProducts $argList"
                            $cmd = "cmd /c cscript //Nologo $cmdLine"
                            Invoke-Expression $cmd
                        }

                    }
                    "Project" {
                        Write-Host "`tRemoving Project products..."
                        WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Removing Project products..." -LogFilePath $LogFilePath
                        $ProjectProductName = $ProjectProduct.Name

                        switch($ProjectProduct.Version){
                            "11" {
                                $ActionFile = "$scriptPath\$03VBS"
                            }
                            "12" {
                                $ActionFile = "$scriptPath\$07VBS"
                            }
                            "14" {
                                $ActionFile = "$scriptPath\$10VBS"
                            }
                            "15" {
                                if(!$isProjectC2R){
                                    $ActionFile = "$scriptPath\$15MSIVBS"
                                } else {
                                    if($RemoveClickToRunVersions){
                                        Remove-OfficeClickToRun -C2RProductsToRemove "ProjectProXVolume", "ProjectStdXVolume","ProjectProRetail"
                                    } else {
                                        WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Project cannot be removed. Use the -RemoveClickToRunVersions parameter to remove Click-To-Run 2016 installs." -LogFilePath $LogFilePath
                                        throw "Project cannot be removed. Use the -RemoveClickToRunVersions parameter to remove Click-To-Run 2016 installs."
                                    }
                                }
                            }
                            "16" {
                                if($Remove2016Installs){
                                    if(!$isProjectC2R){
                                        $ActionFile = "$scriptPath\$16MSIVBS"
                                    } else {
                                        if($RemoveClickToRunVersions){
                                            Remove-OfficeClickToRun -C2RProductsToRemove "ProjectProXVolume", "ProjectStdXVolume","ProjectProRetail"
                                        } else {
                                            WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Project cannot be removed. Use the -RemoveClickToRunVersions parameter to remove Click-To-Run 2016 installs." -LogFilePath $LogFilePath
                                            throw "Project cannot be removed. Use the -RemoveClickToRunVersions parameter to remove Click-To-Run 2016 installs."
                                        }
                                        }
                                    } else {
                                        WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Project cannot be removed. Use the -RemoveClickToRunVersions and -Remove2016Installs parameters to remove Click-To-Run 2016 installs." -LogFilePath $LogFilePath
                                        throw "Project cannot be removed. Use the -RemoveClickToRunVersions and -Remove2016Installs parameters to remove Click-To-Run 2016 installs."
                                    }
                                }
                            }
                        if($ActionFile -And (Test-Path -Path $ActionFile)){
                            $cmdLine = """$ActionFile"" $ProjectProductName $argList"
                            $cmd = "cmd /c cscript //Nologo $cmdLine"
                            Invoke-Expression $cmd
                        }
                    }
                }
            }
        } else {
            Write-Host "`tRemoving all Office products..."
            WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Removing all Office products..." -LogFilePath $LogFilePath

            foreach($product in $officeVersions){
                try{
                    switch -wildcard ($product.Version){
                        "11.*"{
                            if(!$office03Removed){
                                $ActionFile = "$scriptPath\$03VBS"
                                $cmdLine = """$ActionFile"" CLIENTALL $argList"
                                $cmd = "cmd /c cscript //Nologo $cmdLine"
                                Invoke-Expression $cmd
                                $office03Removed = $true
                            }
                        }
                        "12.*"{
                            if(!$office07Removed){
                                $ActionFile = "$scriptPath\$07VBS"
                                $cmdLine = """$ActionFile"" CLIENTALL $argList"
                                $cmd = "cmd /c cscript //Nologo $cmdLine"
                                Invoke-Expression $cmd
                                $office07Removed = $true
                            }
                        }
                        "14.*"{
                            if(!$office10Removed){
                                $ActionFile = "$scriptPath\$10VBS"
                                $cmdLine = """$ActionFile"" CLIENTALL $argList"
                                $cmd = "cmd /c cscript //Nologo $cmdLine"
                                Invoke-Expression $cmd
                                $office10Removed = $true
                            }
                        }
                        "15.*"{
                            if(!$office15Removed){
                                if(!$c2r2013Installed){
                                    $ActionFile = "$scriptPath\$15MSIVBS"
                                } else {
                                    if($RemoveClickToRunVersions){
                                        $ActionFile = "$scriptPath\$c2rVBS"
                                    } else {
                                        WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Office 2013 cannot be removed if 2013 Click-To-Run is installed. Use the -RemoveClickToRunVersions parameter to remove Click-To-Run installs." -LogFilePath $LogFilePath
                                        throw "Office 2013 cannot be removed if 2013 Click-To-Run is installed. Use the -RemoveClickToRunVersions parameter to remove Click-To-Run installs."
                                    }
                                }

                                $cmdLine = """$ActionFile"" CLIENTALL $argList"
                                $cmd = "cmd /c cscript //Nologo $cmdLine"
                                Invoke-Expression $cmd
                                $office15Removed = $true
                            }
                        }
                        "16.*"{
                            if($Remove2016Installs){
                                if($product.ClickToRun -eq $true){
                                    $c2r2016Installed = $true
                                }

                                if(!$c2r2016Installed){
                                    $ActionFile = "$scriptPath\$16MSIVBS"
                                } else {
                                    if($RemoveClickToRunVersions){
                                        $ActionFile = "$scriptPath\$c2rVBS"  
                                    } else {
                                        WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Office 2016 cannot be removed if 2016 Click-To-Run is installed. Use the -RemoveClickToRunVersions parameter to remove Click-To-Run installs." -LogFilePath $LogFilePath
                                        throw "Office 2016 cannot be removed if 2016 Click-To-Run is installed. Use the -RemoveClickToRunVersions parameter to remove Click-To-Run installs."
                                    }
                                }

                                $cmdLine = """$ActionFile"" CLIENTALL $argList"
                                $cmd = "cmd /c cscript //Nologo $cmdLine"
                                Invoke-Expression $cmd
                                $office16Removed = $true
                            }
                        }
                    }
                } catch {}
            }
        }
    }
  }
}

Function Remove-OfficeClickToRun {
<#
.Synopsis
Removes the Click to Run version of Office installed.

.DESCRIPTION
If Office Click-to-Run is installed the administrator will be prompted to confirm
uninstallation. A configuration file will be generated and used to remove all Office CTR 
products.

.PARAMETER ComputerName
The computer or list of computers from which to query 

.EXAMPLE
Remove-OfficeClickToRun

Description:
Will uninstall Office Click-to-Run.
#>
    [CmdletBinding()]
    Param(
        [string[]] $ComputerName = $env:COMPUTERNAME,

        [string] $RemoveCTRXmlPath = "$env:PUBLIC\Documents\RemoveCTRConfig.xml",

        [Parameter()]
        [bool] $WaitForInstallToFinish = $true,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $TargetFilePath = $NULL,

        [Parameter()]
        [ValidateSet("All","O365ProPlusRetail","O365BusinessRetail","VisioProRetail","ProjectProRetail", "SPDRetail", "VisioProXVolume", "VisioStdXVolume", 
                     "ProjectProXVolume", "ProjectStdXVolume", "InfoPathRetail", "SkypeforBusinessEntryRetail", "LyncEntryRetail")]
        [string[]]$C2RProductsToRemove = "All",

        [Parameter()]
        [string]$LogFilePath
    )

     Process{
        $currentFileName = Get-CurrentFileName
        Set-Alias -name LINENUM -value Get-CurrentLineNumber 

        $scriptRoot = GetScriptRoot

        newCTRRemoveXml | Out-File $RemoveCTRXmlPath
       
        if($C2RProductsToRemove -ne "All"){
            foreach($product in $C2RProductsToRemove){
                #Load the xml
                [System.Xml.XmlDocument]$ConfigFile = New-Object System.Xml.XmlDocument
                $content = Get-Content $RemoveCTRXmlPath
                $ConfigFile.LoadXml($content) | Out-Null

                #Set the values
                $RemoveElement = $ConfigFile.Configuration.Remove

                $isValidProduct = (Get-ODTOfficeProductLanguages | ? {$_.DisplayName -eq $product}).DisplayName

                if($isValidProduct  -ne $NULL){
                    [System.Xml.XmlElement]$ProductElement = $ConfigFile.Configuration.Remove.Product | where {$_.ID -eq $product}
                    if($ProductElement -eq $NULL){
                        [System.Xml.XmlElement]$ProductElement = $ConfigFile.CreateElement("Product")
                        $RemoveElement.appendChild($ProductElement) | Out-Null
                        $ProductElement.SetAttribute("ID", $product) | Out-Null
                    }

                    #Add the languages
                    $LanguageIds = (Get-ODTOfficeProductLanguages -ProductId $product).Languages
                    foreach($LanguageId in $LanguageIds){
                        [System.Xml.XmlElement]$LanguageElement = $ProductElement.Language | Where {$_.ID -eq $LanguageId}
                        if($LanguageElement -eq $NULL){
                            [System.Xml.XmlElement]$LanguageElement = $ConfigFile.CreateElement("Language")
                            $ProductElement.AppendChild($LanguageElement) | Out-Null
                            $LanguageElement.SetAttribute("ID", $LanguageId) | Out-Null
                        }
                    }

                    #Save the XML file
                    $ConfigFile.Save($RemoveCTRXmlPath) | Out-Null
                    $global:saveLastFilePath = $RemoveCTRXmlPath
                }
            }

            $RemoveAllElement = $ConfigFile.Configuration.Remove.All
            if($RemoveAllElement -ne $NULL){
                $ConfigFile.Configuration.Remove.RemoveAttribute("All") | Out-Null
            }

            #Save the XML file
            $ConfigFile.Save($RemoveCTRXmlPath) | Out-Null
            $global:saveLastFilePath = $RemoveCTRXmlPath
        }

        [bool] $isInPipe = $true
        if (($PSCmdlet.MyInvocation.PipelineLength -eq 1) -or ($PSCmdlet.MyInvocation.PipelineLength -eq $PSCmdlet.MyInvocation.PipelinePosition)) {
            $isInPipe = $false
        }
            
        $c2rVersion = Get-OfficeVersion | Where-Object {$_.ClickToRun -eq "True" -and $_.DisplayName -match "Microsoft Office 365"}
        if ( $c2rVersion.Count -gt 0) {
            $c2rVersion =  $c2rVersion[0]
        }

        $c2rName = $c2rVersion.DisplayName
             
        if($c2rVersion) {
            if(!($isInPipe)) {
                Write-Host "Please wait while $c2rName is being uninstalled..."
                WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Please wait while $c2rName is being uninstalled..." -LogFilePath $LogFilePath
            }            
        }
   
        if($c2rVersion.Version -like "15*"){
            $OdtExe = "$scriptRoot\Office2013Setup.exe"
        }
        else{
            $OdtExe = "$scriptRoot\Office2016Setup.exe"
        } 

        
        $cmdLine = '"' + $OdtExe + '"'
        $cmdArgs = "/configure " + '"' + $RemoveCTRXmlPath + '"'

        StartProcess -execFilePath $cmdLine -execParams $cmdArgs -WaitForExit $true 
                        
        [bool] $c2rTest = $false 
        if( Get-OfficeVersion | Where-Object {$_.ClickToRun -eq "True"} ){
            $c2rTest = $true
        }

        if($c2rVersion){
            if(!($c2rTest)){                           
                if (!($isInPipe)) {                        
                    Write-Host "Office Click-to-Run has been successfully uninstalled." 
                    WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Office Click-to-Run has been successfully uninstalled." -LogFilePath $LogFilePath 
                }
            }
        }                                      
                                                                               
        if ($isInPipe) {
            $results = new-object PSObject[] 0;
            $Result = New-Object -TypeName PSObject 
            Add-Member -InputObject $Result -MemberType NoteProperty -Name "TargetFilePath" -Value $TargetFilePath
            $Result
        }
    }
}

Function newCTRRemoveXml {
#Create a xml configuration file to remove all Office CTR products.
@"
<Configuration>
  <Remove All="True">
  </Remove>
  <Display Level="None" AcceptEULA="TRUE" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
</Configuration>
"@
}

Function StartProcess {
	Param
	(
        [Parameter()]
		[String]$execFilePath,

        [Parameter()]
        [String]$execParams,

        [Parameter()]
        [bool]$WaitForExit = $false,

        [Parameter()]
        [string]$LogFilePath
	)

    $currentFileName = Get-CurrentFileName
    Set-Alias -name LINENUM -value Get-CurrentLineNumber 

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
        WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError $_
    }
}

Function IsSupportedLanguage() {
    Param(
           [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
           [string] $Language,

           [Parameter()]
           [bool] $ShowLanguages = $true,

           [Parameter()]
           [string]$LogFilePath
        )
        
        $currentFileName = Get-CurrentFileName
        Set-Alias -name LINENUM -value Get-CurrentLineNumber

        $lang = $validLanguages | where {$_.ToString().ToUpper().EndsWith("|$Language".ToUpper())}
          
        if (!($lang)) {
           if ($ShowLanguages) {
              Write-Host
              Write-Host "Invalid or Unsupported Language. Please select a language." -NoNewLine -BackgroundColor Red
              WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Invalid or Unsupported Language. Please select a language." -LogFilePath $LogFilePath
              Write-Host

              return SelectLanguage 
           } else {
              WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "Invalid or Unsupported Language: $Language" -LogFilePath $LogFilePath
              throw "Invalid or Unsupported Language: $Language"
           }
           
        }
        
        return $Language
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

Function Get-OfficeVersion {
<#
.Synopsis
Gets the Office Version installed on the computer
.DESCRIPTION
This function will query the local or a remote computer and return the information about Office Products installed on the computer
.NOTES   
Name: Get-OfficeVersion
Version: 1.0.5
DateCreated: 2015-07-01
DateUpdated: 2016-10-14
.LINK
https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts
.PARAMETER ComputerName
The computer or list of computers from which to query 
.PARAMETER ShowAllInstalledProducts
Will expand the output to include all installed Office products
.EXAMPLE
Get-OfficeVersion
Description:
Will return the locally installed Office product
.EXAMPLE
Get-OfficeVersion -ComputerName client01,client02
Description:
Will return the installed Office product on the remote computers
.EXAMPLE
Get-OfficeVersion | select *
Description:
Will return the locally installed Office product with all of the available properties
#>
[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true, Position=0)]
    [string[]]$ComputerName = $env:COMPUTERNAME,
    [switch]$ShowAllInstalledProducts,
    [System.Management.Automation.PSCredential]$Credentials
)

begin {
    $HKLM = [UInt32] "0x80000002"
    $HKCR = [UInt32] "0x80000000"

    $excelKeyPath = "Excel\DefaultIcon"
    $wordKeyPath = "Word\DefaultIcon"
   
    $installKeys = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
                   'SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall'

    $officeKeys = 'SOFTWARE\Microsoft\Office',
                  'SOFTWARE\Wow6432Node\Microsoft\Office'

    $defaultDisplaySet = 'DisplayName','Version', 'ComputerName'

    $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet',[string[]]$defaultDisplaySet)
    $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
}

process {

 $results = new-object PSObject[] 0;
 $MSexceptionList = "mui","visio","project","proofing","visual"

 foreach ($computer in $ComputerName) {
    if ($Credentials) {
       $os=Get-WMIObject win32_operatingsystem -computername $computer -Credential $Credentials
    } else {
       $os=Get-WMIObject win32_operatingsystem -computername $computer
    }

    $osArchitecture = $os.OSArchitecture

    if ($Credentials) {
       $regProv = Get-Wmiobject -list "StdRegProv" -namespace root\default -computername $computer -Credential $Credentials
    } else {
       $regProv = Get-Wmiobject -list "StdRegProv" -namespace root\default -computername $computer
    }

    [System.Collections.ArrayList]$VersionList = New-Object -TypeName System.Collections.ArrayList
    [System.Collections.ArrayList]$PathList = New-Object -TypeName System.Collections.ArrayList
    [System.Collections.ArrayList]$PackageList = New-Object -TypeName System.Collections.ArrayList
    [System.Collections.ArrayList]$ClickToRunPathList = New-Object -TypeName System.Collections.ArrayList
    [System.Collections.ArrayList]$ConfigItemList = New-Object -TypeName  System.Collections.ArrayList
    $ClickToRunList = new-object PSObject[] 0;

    foreach ($regKey in $officeKeys) {
       $officeVersion = $regProv.EnumKey($HKLM, $regKey)
       foreach ($key in $officeVersion.sNames) {
          if ($key -match "\d{2}\.\d") {
            if (!$VersionList.Contains($key)) {
              $AddItem = $VersionList.Add($key)
            }

            $path = join-path $regKey $key

            $configPath = join-path $path "Common\Config"
            $configItems = $regProv.EnumKey($HKLM, $configPath)
            if ($configItems) {
               foreach ($configId in $configItems.sNames) {
                 if ($configId) {
                    $Add = $ConfigItemList.Add($configId.ToUpper())
                 }
               }
            }

            $cltr = New-Object -TypeName PSObject
            $cltr | Add-Member -MemberType NoteProperty -Name InstallPath -Value ""
            $cltr | Add-Member -MemberType NoteProperty -Name UpdatesEnabled -Value $false
            $cltr | Add-Member -MemberType NoteProperty -Name UpdateUrl -Value ""
            $cltr | Add-Member -MemberType NoteProperty -Name StreamingFinished -Value $false
            $cltr | Add-Member -MemberType NoteProperty -Name Platform -Value ""
            $cltr | Add-Member -MemberType NoteProperty -Name ClientCulture -Value ""
            
            $packagePath = join-path $path "Common\InstalledPackages"
            $clickToRunPath = join-path $path "ClickToRun\Configuration"
            $virtualInstallPath = $regProv.GetStringValue($HKLM, $clickToRunPath, "InstallationPath").sValue

            [string]$officeLangResourcePath = join-path  $path "Common\LanguageResources"
            $mainLangId = $regProv.GetDWORDValue($HKLM, $officeLangResourcePath, "SKULanguage").uValue
            if ($mainLangId) {
                $mainlangCulture = [globalization.cultureinfo]::GetCultures("allCultures") | where {$_.LCID -eq $mainLangId}
                if ($mainlangCulture) {
                    $cltr.ClientCulture = $mainlangCulture.Name
                }
            }

            [string]$officeLangPath = join-path  $path "Common\LanguageResources\InstalledUIs"
            $langValues = $regProv.EnumValues($HKLM, $officeLangPath);
            if ($langValues) {
               foreach ($langValue in $langValues) {
                  $langCulture = [globalization.cultureinfo]::GetCultures("allCultures") | where {$_.LCID -eq $langValue}
               } 
            }

            if ($virtualInstallPath) {

            } else {
              $clickToRunPath = join-path $regKey "ClickToRun\Configuration"
              $virtualInstallPath = $regProv.GetStringValue($HKLM, $clickToRunPath, "InstallationPath").sValue
            }

            if ($virtualInstallPath) {
               if (!$ClickToRunPathList.Contains($virtualInstallPath.ToUpper())) {
                  $AddItem = $ClickToRunPathList.Add($virtualInstallPath.ToUpper())
               }

               $cltr.InstallPath = $virtualInstallPath
               $cltr.StreamingFinished = $regProv.GetStringValue($HKLM, $clickToRunPath, "StreamingFinished").sValue
               $cltr.UpdatesEnabled = $regProv.GetStringValue($HKLM, $clickToRunPath, "UpdatesEnabled").sValue
               $cltr.UpdateUrl = $regProv.GetStringValue($HKLM, $clickToRunPath, "UpdateUrl").sValue
               $cltr.Platform = $regProv.GetStringValue($HKLM, $clickToRunPath, "Platform").sValue
               $cltr.ClientCulture = $regProv.GetStringValue($HKLM, $clickToRunPath, "ClientCulture").sValue
               $ClickToRunList += $cltr
            }

            $packageItems = $regProv.EnumKey($HKLM, $packagePath)
            $officeItems = $regProv.EnumKey($HKLM, $path)

            foreach ($itemKey in $officeItems.sNames) {
              $itemPath = join-path $path $itemKey
              $installRootPath = join-path $itemPath "InstallRoot"

              $filePath = $regProv.GetStringValue($HKLM, $installRootPath, "Path").sValue
              if (!$PathList.Contains($filePath)) {
                  $AddItem = $PathList.Add($filePath)
              }
            }

            foreach ($packageGuid in $packageItems.sNames) {
              $packageItemPath = join-path $packagePath $packageGuid
              $packageName = $regProv.GetStringValue($HKLM, $packageItemPath, "").sValue
            
              if (!$PackageList.Contains($packageName)) {
                if ($packageName) {
                   $AddItem = $PackageList.Add($packageName.Replace(' ', '').ToLower())
                }
              }
            }

          }
       }
    }

    

    foreach ($regKey in $installKeys) {
        $keyList = new-object System.Collections.ArrayList
        $keys = $regProv.EnumKey($HKLM, $regKey)

        foreach ($key in $keys.sNames) {
           $path = join-path $regKey $key
           $installPath = $regProv.GetStringValue($HKLM, $path, "InstallLocation").sValue
           if (!($installPath)) { continue }
           if ($installPath.Length -eq 0) { continue }

           $buildType = "64-Bit"
           if ($osArchitecture -eq "32-bit") {
              $buildType = "32-Bit"
           }

           if ($regKey.ToUpper().Contains("Wow6432Node".ToUpper())) {
              $buildType = "32-Bit"
           }

           if ($key -match "{.{8}-.{4}-.{4}-1000-0000000FF1CE}") {
              $buildType = "64-Bit" 
           }

           if ($key -match "{.{8}-.{4}-.{4}-0000-0000000FF1CE}") {
              $buildType = "32-Bit" 
           }

           if ($modifyPath) {
               if ($modifyPath.ToLower().Contains("platform=x86")) {
                  $buildType = "32-Bit"
               }

               if ($modifyPath.ToLower().Contains("platform=x64")) {
                  $buildType = "64-Bit"
               }
           }

           $primaryOfficeProduct = $false
           $officeProduct = $false
           foreach ($officeInstallPath in $PathList) {
             if ($officeInstallPath) {
                try{
                $installReg = "^" + $installPath.Replace('\', '\\')
                $installReg = $installReg.Replace('(', '\(')
                $installReg = $installReg.Replace(')', '\)')
                if ($officeInstallPath -match $installReg) { $officeProduct = $true }
                } catch {}
             }
           }

           if (!$officeProduct) { continue };
           
           $name = $regProv.GetStringValue($HKLM, $path, "DisplayName").sValue          

           $primaryOfficeProduct = $true
           if ($ConfigItemList.Contains($key.ToUpper()) -and $name.ToUpper().Contains("MICROSOFT OFFICE")) {
              foreach($exception in $MSexceptionList){
                 if($name.ToLower() -match $exception.ToLower()){
                    $primaryOfficeProduct = $false
                 }
              }
           } else {
              $primaryOfficeProduct = $false
           }

           $clickToRunComponent = $regProv.GetDWORDValue($HKLM, $path, "ClickToRunComponent").uValue
           $uninstallString = $regProv.GetStringValue($HKLM, $path, "UninstallString").sValue
           if (!($clickToRunComponent)) {
              if ($uninstallString) {
                 if ($uninstallString.Contains("OfficeClickToRun")) {
                     $clickToRunComponent = $true
                 }
              }
           }

           $modifyPath = $regProv.GetStringValue($HKLM, $path, "ModifyPath").sValue 
           $version = $regProv.GetStringValue($HKLM, $path, "DisplayVersion").sValue

           $cltrUpdatedEnabled = $NULL
           $cltrUpdateUrl = $NULL
           $clientCulture = $NULL;

           [string]$clickToRun = $false

           if ($clickToRunComponent) {
               $clickToRun = $true
               if ($name.ToUpper().Contains("MICROSOFT OFFICE")) {
                  $primaryOfficeProduct = $true
               }

               foreach ($cltr in $ClickToRunList) {
                 if ($cltr.InstallPath) {
                   if ($cltr.InstallPath.ToUpper() -eq $installPath.ToUpper()) {
                       $cltrUpdatedEnabled = $cltr.UpdatesEnabled
                       $cltrUpdateUrl = $cltr.UpdateUrl
                       if ($cltr.Platform -eq 'x64') {
                           $buildType = "64-Bit" 
                       }
                       if ($cltr.Platform -eq 'x86') {
                           $buildType = "32-Bit" 
                       }
                       $clientCulture = $cltr.ClientCulture
                   }
                 }
               }
           }
           
           if (!$primaryOfficeProduct) {
              if (!$ShowAllInstalledProducts) {
                  continue
              }
           }

           $object = New-Object PSObject -Property @{DisplayName = $name; Version = $version; InstallPath = $installPath; ClickToRun = $clickToRun; 
                     Bitness=$buildType; ComputerName=$computer; ClickToRunUpdatesEnabled=$cltrUpdatedEnabled; ClickToRunUpdateUrl=$cltrUpdateUrl;
                     ClientCulture=$clientCulture }
           $object | Add-Member MemberSet PSStandardMembers $PSStandardMembers
           $results += $object

        }
    }

  }

  $results = Get-Unique -InputObject $results 

  return $results;
}

}

Function IsDotSourced() {
  [CmdletBinding(SupportsShouldProcess=$true)]
  param(
    [Parameter(ValueFromPipelineByPropertyName=$true)]
    [string]$InvocationLine = ""
  )
  $cmdLine = $InvocationLine.Trim()
  Do {
    $cmdLine = $cmdLine.Replace(" ", "")
  } while($cmdLine.Contains(" "))

  $dotSourced = $false
  if ($cmdLine -match '^\.\\') {
     $dotSourced = $false
  } else {
     $dotSourced = ($cmdLine -match '^\.')
  }

  return $dotSourced
}

Function GetScriptRoot() {
 process {
     [string]$scriptPath = "."

     if ($PSScriptRoot) {
       $scriptPath = $PSScriptRoot
     } else {
       $scriptPath = (Get-Item -Path ".\").FullName
     }

     return $scriptPath
 }
}

function GetProductName {
param(
    [Parameter()]
    [string]$ProductName,

    [Parameter()]
    [string]$LogFilePath
)
    $defaultDisplaySet = 'DisplayName','Name','Version'
    $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet',[string[]]$defaultDisplaySet)
    $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
    $results = New-Object PSObject[] 0;
    
    $currentFileName = Get-CurrentFileName
    Set-Alias -name LINENUM -value Get-CurrentLineNumber

    if($ProductName -eq 'MainOfficeProduct'){
        $MainOfficeProducts = @()
        #$Products = (Get-OfficeVersion).DisplayName | select -Unique
        $MainOfficeProducts = (Get-OfficeVersion)
        if($MainOfficeProducts.GetType().Name -eq "Object[]"){
            $primaryOfficeLanguage = GetClientCulture
            $MainOfficeProduct = (Get-OfficeVersion) | ? {$_.DisplayName -match $primaryOfficeLanguage}
            $ProductName = $MainOfficeProduct.DisplayName
        } else {
            $ProductName = $MainOfficeProducts.DisplayName
        }
    }
    
    WriteToLogFile -LNumber $(LINENUM) -FName $currentFileName -ActionError "ProductName set to $ProductName" -LogFilePath $LogFilePath 
        
    $HKLM = [UInt32] "0x80000002"
    $HKCR = [UInt32] "0x80000000"
 
    $installKeys = 'SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall',
                   'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
                   

    $regProv = Get-WmiObject -list "StdRegProv" -namespace root\default -ComputerName $env:COMPUTERNAME

    if($ProductName.ToLower() -match "visio" -or $ProductName.ToLower() -match "project"){
        $ProductName = " " + $ProductName + " "
    }

    foreach ($regKey in $installKeys) {
        $keyList = New-Object System.Collections.ArrayList
        $keys = $regProv.EnumKey($HKLM, $regKey)

        foreach ($key in $keys.sNames) {
            $path = Join-Path $regKey $key
            $name = $regProv.GetStringValue($HKLM, $path, "DisplayName").sValue
            $version = $regProv.GetStringValue($HKLM, $path, "DisplayVersion").sValue
            
            if($name){
                if($name.ToLower() -match $ProductName.ToLower()){
                    if($path -notmatch "{.{8}-.{4}-.{4}-.{4}-0000000FF1CE}"){
                        if($name -match "Language Pack"){
                            if($key.Split(".")[1] -ne $null){
                                $regex = "^[^.]*"
                                $string = $key -replace $regex, ""
                                $prodName = $string.trim(".")
                            }
                        } else {
                            if($key.Split(".")[1] -ne $null){
                                $prodName = $key.Split(".")[1]
                            } else {
                                $prodName = $key
                            }
                        }
                        $prodVersion = $version.Split(".")[0]
                        $DisplayName = $name

                        $object = New-Object PSObject -Property @{DisplayName = $DisplayName; Name = $prodName; Version = $prodVersion }
                        $object | Add-Member MemberSet PSStandardMembers $PSStandardMembers
                        $results += $object
                    }
                }
            }
        }
    }

    return $Results

}

function Get-ODTOfficeProductLanguages {
    Param(
        [Parameter()]
        [string]$ComputerName = $env:COMPUTERNAME,

        [Parameter()]
        [string]$ProductId
    )

    Begin {
        $defaultDisplaySet = 'DisplayName','Languages'
        $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet(‘DefaultDisplayPropertySet’,[string[]]$defaultDisplaySet)
        $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
        $results = New-Object PSObject[] 0
    }

    Process {
        $regProv = Get-Wmiobject -list "StdRegProv" -namespace root\default -ComputerName $ComputerName -ErrorAction Stop
        $officeConfig = getCTRConfig -regProv $regProv
        [System.XML.XMLDocument]$ConfigFile = New-Object System.XML.XMLDocument

        if(!$ProductId){
            $productReleaseIds = $officeConfig.ProductReleaseIds
            $splitProducts = $productReleaseIds.Split(',')
        } else {
            $splitProducts = $ProductId
        }

        foreach($product in $splitProducts){
            $officeAddLangs = odtGetOfficeLanguages -ConfigDoc $ConfigFile -OfficeKeyPath $officeConfig.OfficeKeyPath -ProductId $product

            $object = New-Object PSObject -Property @{DisplayName = $product; Languages = $officeAddLangs}
            $object | Add-Member MemberSet PSStandardMembers $PSStandardMembers
            $results += $object
        }
  
        return $results
    }
}

function getCTRConfig() {
    param(
       [Parameter(ValueFromPipelineByPropertyName=$true)]
       $regProv = $NULL
    )

    $HKLM = [UInt32] "0x80000002"
    $computerName = $env:COMPUTERNAME

    if (!($regProv)) {
        $regProv = Get-Wmiobject -list "StdRegProv" -namespace root\default -computername $computerName -ErrorAction Stop
    }
    
    $officeCTRKeys = 'SOFTWARE\Microsoft\Office\15.0\ClickToRun',
                     'SOFTWARE\Wow6432Node\Microsoft\Office\15.0\ClickToRun',
                     'SOFTWARE\Microsoft\Office\ClickToRun',
                     'SOFTWARE\Wow6432Node\Microsoft\Office\ClickToRun'

    $Object = New-Object PSObject
    $Object | Add-Member Noteproperty ClickToRunInstalled $false

    [string]$officeKeyPath = "";
    foreach ($regPath in $officeCTRKeys) {
       [string]$installPath = $regProv.GetStringValue($HKLM, $regPath, "InstallPath").sValue
       if ($installPath) {
          if ($installPath.Length -gt 0) {
              $officeKeyPath = $regPath;
              break;
          }
       }
    }

    if ($officeKeyPath.Length -gt 0) {
        $Object.ClickToRunInstalled = $true

        $configurationPath = join-path $officeKeyPath "Configuration"

        [string]$platform = $regProv.GetStringValue($HKLM, $configurationPath, "Platform").sValue
        [string]$clientCulture = $regProv.GetStringValue($HKLM, $configurationPath, "ClientCulture").sValue
        [string]$productIds = $regProv.GetStringValue($HKLM, $configurationPath, "ProductReleaseIds").sValue
        [string]$versionToReport = $regProv.GetStringValue($HKLM, $configurationPath, "VersionToReport").sValue
        [string]$updatesEnabled = $regProv.GetStringValue($HKLM, $configurationPath, "UpdatesEnabled").sValue
        [string]$updateUrl = $regProv.GetStringValue($HKLM, $configurationPath, "UpdateUrl").sValue
        [string]$updateDeadline = $regProv.GetStringValue($HKLM, $configurationPath, "UpdateDeadline").sValue

        if (!($productIds)) {
            $productIds = ""
            $officeActivePath = Join-Path $officeKeyPath "ProductReleaseIDs\Active"
            $officeProducts = $regProv.EnumKey($HKLM, $officeActivePath)

            foreach ($productName in $officeProducts.sNames) {
               if ($productName.ToLower() -eq "stream") { continue }
               if ($productName.ToLower() -eq "culture") { continue }
               if ($productIds.Length -gt 0) { $productIds += "," }
               $productIds += "$productName"
            }
        }

        $splitProducts = $productIds.Split(',');

        if ($platform.ToLower() -eq "x86") {
            $platform = "32"
        } else {
            $platform = "64"
        }

        $Object | add-member Noteproperty Platform $platform
        $Object | add-member Noteproperty ClientCulture $clientCulture
        $Object | add-member Noteproperty ProductReleaseIds $productIds
        $Object | add-member Noteproperty Version $versionToReport
        $Object | add-member Noteproperty UpdatesEnabled $updatesEnabled
        $Object | add-member Noteproperty UpdateUrl $updateUrl
        $Object | add-member Noteproperty UpdateDeadline $updateDeadline
        $Object | add-member Noteproperty OfficeKeyPath $officeKeyPath
        
    } 

    return $Object 

}

function odtGetOfficeLanguages() {
    param(
       [Parameter(ValueFromPipelineByPropertyName=$true)]
       [System.XML.XMLDocument]$ConfigDoc = $NULL,
              
       [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
       [string]$OfficeKeyPath = $NULL,

       [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
       [string]$ProductId = $NULL
    )

    begin {
        $HKLM = [UInt32] "0x80000002"
        $HKCR = [UInt32] "0x80000000"
    }

    process {
        [System.Collections.ArrayList]$appLanguages1 = New-Object System.Collections.ArrayList

        #SOFTWARE\Wow6432Node\Microsoft\Office\14.0\Common\LanguageResources\InstalledUIs

        $productsPath = join-path $officeKeyPath "ProductReleaseIDs\Active\$ProductId"
        $installedCultures = $regProv.EnumKey($HKLM, $productsPath)
        
        foreach ($installedCulture in $installedCultures.sNames) {
        if($installedCulture){
            if ($installedCulture.Contains("-") -and !($installedCulture.ToLower() -eq "x-none")) {
                $addItem = $appLanguages1.Add($installedCulture) 
            }
            }
        }

        if ($appLanguages1.Count) {
            $productsPath = join-path $officeKeyPath "ProductReleaseIDs\Active\$ProductId"
        } else {
            $productReleasePath = Join-Path $officeKeyPath "ProductReleaseIDs"
            $guids= $regProv.EnumKey($HKLM, $productReleasePath)

            foreach ($guid in $guids.sNames) {

                $productsPath = Join-Path $officeKeyPath "ProductReleaseIDs\$guid\$ProductId.16"
                $installedCultures = $regProv.EnumKey($HKLM, $productsPath)
      
                foreach ($installedCulture in $installedCultures.sNames) {
                   if($installedCulture){
                      if ($installedCulture.Contains("-") -and !($installedCulture.ToLower() -eq "x-none")) {
                            $addItem = $appLanguages1.Add($installedCulture) 
                      }
                   }
                }

                if ($appLanguages1.Count) {
                   $productsPath = Join-Path $officeKeyPath "ProductReleaseIDs\Culture\$ProductId"
                }
 
            }
        }

        return $appLanguages1;
    }
}

function GetClientCulture{
    Param(
        [string]$computer = $env:COMPUTERNAME
    )
    
    $HKLM = [UInt32] "0x80000002"

    $officeKeys = 'SOFTWARE\Microsoft\Office',
                  'SOFTWARE\Wow6432Node\Microsoft\Office'

    $regProv = Get-WmiObject -List "StdRegProv" -Namespace root\default -ComputerName $computer

    foreach ($regKey in $officeKeys) {
        $officeVersion = $regProv.EnumKey($HKLM, $regKey)
        foreach ($key in $officeVersion.sNames) {
            if($key -match "\d{2}\.\d") {
                $path = join-path $regKey $key
                $clickToRunPath = join-path $path "ClickToRun\Configuration"
                if(Test-Path "HKLM:\$clickToRunPath"){           
                    $clientCulture = $regProv.GetStringValue($HKLM, $clickToRunPath, "ClientCulture").sValue
                }               
            }
        }
    }

    return $clientCulture
}

function Get-CurrentLineNumber {
    $MyInvocation.ScriptLineNumber
}

function Get-CurrentFileName{
    $MyInvocation.ScriptName.Substring($MyInvocation.ScriptName.LastIndexOf("\")+1)
}

Function WriteToLogFile() {
    param( 
        [Parameter(Mandatory=$true)]
        [string]$LNumber,

        [Parameter(Mandatory=$true)]
        [string]$FName,

        [Parameter(Mandatory=$true)]
        [string]$ActionError,

        [Parameter()]
        [string]$LogFilePath
    )

    try{
        $headerString = "Time".PadRight(30, ' ') + "Line Number".PadRight(15,' ') + "FileName".PadRight(60,' ') + "Action"
        $stringToWrite = $(Get-Date -Format G).PadRight(30, ' ') + $($LNumber).PadRight(15, ' ') + $($FName).PadRight(60,' ') + $ActionError

        if(!$LogFilePath){
            $LogFilePath = "$env:windir\Temp\" + (Get-Date -Format u).Substring(0,10)+"_OfficeDeploymentLog.txt"
        }
        if(Test-Path $LogFilePath){
             Add-Content $LogFilePath $stringToWrite
        }
        else{#if not exists, create new
             Add-Content $LogFilePath $headerString
             Add-Content $LogFilePath $stringToWrite
        }
    } catch [Exception]{
        Write-Host $_
    }
}

$dotSourced = IsDotSourced -InvocationLine $MyInvocation.Line

if (!($dotSourced)) {
   Remove-PreviousOfficeInstalls -RemoveClickToRunVersions $RemoveClickToRunVersions -Remove2016Installs $Remove2016Installs -Force $Force -KeepUserSettings $KeepUserSettings -KeepLync $KeepLync -NoReboot $NoReboot -ProductsToRemove $ProductsToRemove -LogFilePath $LogFilePath
}