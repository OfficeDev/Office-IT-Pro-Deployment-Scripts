﻿[CmdletBinding(SupportsShouldProcess=$true)]
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

[Parameter()]
[ValidateSet("AllOfficeProducts","MainOfficeProduct","Visio","Project")]
[string[]]$ProductsToRemove
)

Function Remove-PreviousOfficeInstalls{
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
    [string[]]$ProductsToRemove
  )

  Process {
    $c2rVBS = "OffScrubc2r.vbs"
    $03VBS = "OffScrub03.vbs"
    $07VBS = "OffScrub07.vbs"
    $10VBS = "OffScrub10.vbs"
    $15MSIVBS = "OffScrub_O15msi.vbs"
    $16MSIVBS = "OffScrub_O16msi.vbs"

    $argList = ""
    $argListProducts = @()

    if(!$ProductsToRemove){
        $ProductsToRemove = 'AllOfficeProducts'
    }
    
    if($ProductsToRemove -eq 'AllOfficeProducts'){
        $argListProducts += "CLIENTALL"
    } else {       
        foreach($product in $ProductsToRemove){
            switch($product){
                "MainOfficeProduct"{
                    $MainOfficeProduct = GetProductName -ProductName MainOfficeProduct
                    $argListProducts += $MainOfficeProduct.Name
                }
                "Visio" {
                    $VisioProduct = GetProductName -ProductName Visio
                    $argListProducts += $VisioProduct.Name
                }
                "Project" {
                    $ProjectProduct = GetProductName -ProductName Project
                    $argListProducts += $ProjectProduct.Name
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

    $officeVersions = Get-OfficeVersion -ShowAllInstalledProducts | select *
    $ActionFiles = @()
    
    $removeOffice = $true
    if (!( $officeVersions)) {
       Write-Host "Microsoft Office is not installed"
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

        foreach ($officeVersion in $officeVersions) {
           if($officeVersion.ClicktoRun.ToLower() -eq "true"){
              if ($officeVersion.Version -like '15.*') {
                  $c2r2013Installed = $true
              }
              if ($officeVersion.Version -like '16.*') {
                  $c2r2016Installed = $true
              }
           }
        }

        foreach ($officeVersion in $officeVersions) {
            if($officeVersion.ClicktoRun.ToLower() -eq "true"){
              $removeC2R = $false

              if (!($officeC2RRemoved)) {
                  if ($RemoveClickToRunVersions -and (!($c2r2016Installed))) {
                     $removeC2R = $true
                  }
                  if ($Remove2016Installs -and $RemoveClickToRunVersions) {
                     $removeC2R = $true
                  }
              }

              if ($removeC2R) {
                  Write-Host "`tRemoving Office Click-To-Run..."
                  $ActionFile = "$scriptPath\$c2rVBS"
                  $cmdLine = """$ActionFile"" $argList"
                 
                  if (Test-Path -Path $ActionFile) {
                    $cmd = "cmd /c cscript //Nologo $cmdLine"
                    Invoke-Expression $cmd
                    $officeC2RRemoved = $true
                    $c2r2013Installed = $false
                  } else {
                    throw "Required file missing: $ActionFile"
                  }
                  Write-Host ""
              }

            }
        }

        foreach ($officeVersion in $officeVersions) {
            if($officeVersion.ClicktoRun.ToLower() -ne "true"){
                #Set script file based on office version, if no office detected continue to next computer skipping this one.
                switch -wildcard ($officeVersion.Version)
                {
                    "11.*"
                    {
                        if (!($office03Removed)) {
                            $ActionFile = "$scriptPath\$03VBS"
                            if (Test-Path -Path $ActionFile) {
                                if($ProductsToRemove -ne 'AllOfficeProducts'){
                                    foreach($product in $ProductsToRemove){
                                        switch($product){
                                            "MainOfficeProduct" {
                                                $ProductName = (Get-OfficeVersion).DisplayName | select -Unique
                                                if($ProductName -eq $MainOfficeProduct.DisplayName -and $MainOfficeProduct.DisplayName -match '2003'){
                                                    Write-Host "`n`tRemoving "$MainOfficeProduct.DisplayName"..."
                                                    $MainOfficeProductName = $MainOfficeProduct.Name
                                                    $cmdLine = """$ActionFile"" $MainOfficeProductName $argList"
                                                    $cmd = "cmd /c cscript //Nologo $cmdLine"
                                                    Invoke-Expression $cmd
                                                }
                                            }
                                            "Visio" {
                                                $ProductName = Get-OfficeVersion -ShowAllInstalledProducts | ? {$_.DisplayName -match $product}
                                                foreach($prod in $ProductName){
                                                    if($VisioProduct.DisplayName -eq $prod.DisplayName -and $prod.DisplayName -match '2003'){
                                                        Write-Host "`n`tRemoving "$VisioProduct.DisplayName"..."
                                                        $VisioProductName = $VisioProduct.Name
                                                        $cmdLine = """$ActionFile"" $VisioProductName $argList"
                                                        $cmd = "cmd /c cscript //Nologo $cmdLine"
                                                        Invoke-Expression $cmd
                                                    }
                                                }
                                            }
                                            "Project" {
                                                $ProductName = Get-OfficeVersion -ShowAllInstalledProducts | ? {$_.DisplayName -match $product}
                                                foreach($prod in $ProductName){
                                                    if($ProjectProduct.DisplayName -eq $prod.DisplayName -and $prod.DisplayName -match '2003'){
                                                        Write-Host "`n`tRemoving "$ProjectProduct.DisplayName"..."
                                                        $ProjectProductName = $ProjectProduct.Name
                                                        $cmdLine = """$ActionFile"" $ProjectProductName $argList"
                                                        $cmd = "cmd /c cscript //Nologo $cmdLine"
                                                        Invoke-Expression $cmd
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    Write-Host "`n"
                                    $office03Removed = $true
                                } else {
                                    Write-Host "`n`tRemoving all Office 2003 products..."
                                    $cmdLine = """$ActionFile"" CLIENTALL $argList"
                                    $cmd = "cmd /c cscript //Nologo $cmdLine"
                                    Invoke-Expression $cmd
                                    $office03Removed = $true
                                }
                            } else {
                               throw "Required file missing: $ActionFile"
                            }
                        }
                    }
                    "12.*"
                    {
                        if (!($office07Removed)) {
                            $ActionFile = "$scriptPath\$07VBS"
                            if (Test-Path -Path $ActionFile) {
                                if($ProductsToRemove -ne 'AllOfficeProducts'){
                                    foreach($product in $ProductsToRemove){
                                        switch($product){
                                            "MainOfficeProduct" {
                                                $ProductName = (Get-OfficeVersion).DisplayName | select -Unique
                                                if($ProductName -eq $MainOfficeProduct.DisplayName -and $MainOfficeProduct.DisplayName -match '2007'){
                                                    Write-Host "`n`tRemoving "$MainOfficeProduct.DisplayName"..."
                                                    $MainOfficeProductName = $MainOfficeProduct.Name
                                                    $cmdLine = """$ActionFile"" $MainOfficeProductName $argList"
                                                    $cmd = "cmd /c cscript //Nologo $cmdLine"
                                                    Invoke-Expression $cmd
                                                }
                                            }
                                            "Visio" {
                                                $ProductName = Get-OfficeVersion -ShowAllInstalledProducts | ? {$_.DisplayName -match $product}
                                                foreach($prod in $ProductName){
                                                    if($VisioProduct.DisplayName -eq $prod.DisplayName -and $prod.DisplayName -match '2007'){
                                                        Write-Host "`n`tRemoving "$VisioProduct.DisplayName"..."
                                                        $VisioProductName = $VisioProduct.Name
                                                        $cmdLine = """$ActionFile"" $VisioProductName $argList"
                                                        $cmd = "cmd /c cscript //Nologo $cmdLine"
                                                        Invoke-Expression $cmd
                                                    }
                                                }       
                                            }
                                            "Project" {
                                                $ProductName = Get-OfficeVersion -ShowAllInstalledProducts | ? {$_.DisplayName -match $product}
                                                foreach($prod in $ProductName){
                                                    if($ProjectProduct.DisplayName -eq $prod.DisplayName -and $prod.DisplayName -match '2007'){
                                                        Write-Host "`n`tRemoving "$ProjectProduct.DisplayName"..."
                                                        $ProjectProductName = $ProjectProduct.Name
                                                        $cmdLine = """$ActionFile"" $ProjectProductName $argList"
                                                        $cmd = "cmd /c cscript //Nologo $cmdLine"
                                                        Invoke-Expression $cmd
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    Write-Host "`n"
                                    $office07Removed = $true
                                } else {
                                    Write-Host "`n`tRemoving all Office 2007 products..."
                                    $cmdLine = """$ActionFile"" CLIENTALL $argList"
                                    $cmd = "cmd /c cscript //Nologo $cmdLine"
                                    Invoke-Expression $cmd
                                    $office07Removed = $true
                                }
                            } else {
                               throw "Required file missing: $ActionFile"
                            }                      
                        } 
                    }
                    "14.*"
                    {
                        if (!($office10Removed)) {
                            $ActionFile = "$scriptPath\$10VBS"
                            if (Test-Path -Path $ActionFile) {
                                if($ProductsToRemove -ne 'AllOfficeProducts'){
                                    foreach($product in $ProductsToRemove){
                                        switch($product){
                                            "MainOfficeProduct" {
                                                $ProductName = (Get-OfficeVersion).DisplayName | select -Unique
                                                if($ProductName -eq $MainOfficeProduct.DisplayName -and $MainOfficeProduct.DisplayName -match '2010'){
                                                    Write-Host "`n`tRemoving "$MainOfficeProduct.DisplayName"..."
                                                    $MainOfficeProductName = $MainOfficeProduct.Name
                                                    $cmdLine = """$ActionFile"" $MainOfficeProductName $argList"
                                                    $cmd = "cmd /c cscript //Nologo $cmdLine"
                                                    Invoke-Expression $cmd
                                                }
                                            }
                                            "Visio" {
                                                $ProductName = Get-OfficeVersion -ShowAllInstalledProducts | ? {$_.DisplayName -match $product}
                                                foreach($prod in $ProductName){
                                                    if($VisioProduct.DisplayName -eq $prod.DisplayName -and $prod.DisplayName -match '2010'){
                                                        Write-Host "`n`tRemoving "$VisioProduct.DisplayName"..."
                                                        $VisioProductName = $VisioProduct.Name
                                                        $cmdLine = """$ActionFile"" $VisioProductName $argList"
                                                        $cmd = "cmd /c cscript //Nologo $cmdLine"
                                                        Invoke-Expression $cmd
                                                    }
                                                }
                                            }
                                            "Project" {
                                                $ProductName = Get-OfficeVersion -ShowAllInstalledProducts | ? {$_.DisplayName -match $product}
                                                foreach($prod in $ProductName){
                                                    if($ProjectProduct.DisplayName -eq $prod.DisplayName -and $prod.DisplayName -match '2010'){
                                                        Write-Host "`n`tRemoving "$ProjectProduct.DisplayName"..."
                                                        $ProjectProductName = $ProjectProduct.Name
                                                        $cmdLine = """$ActionFile"" $ProjectProductName $argList"
                                                        $cmd = "cmd /c cscript //Nologo $cmdLine"
                                                        Invoke-Expression $cmd
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    Write-Host "`n"
                                    $office10Removed = $true
                                } else {
                                    Write-Host "`n`tRemoving all Office 2010 products..."
                                    $cmdLine = """$ActionFile"" CLIENTALL $argList"
                                    $cmd = "cmd /c cscript //Nologo $cmdLine"
                                    Invoke-Expression $cmd
                                    $office10Removed = $true
                                }
                            } else {
                               throw "Required file missing: $ActionFile"
                            }                                                     
                        }
                    }
                    "15.*"
                    {
                        if (!($office15Removed)) {
                            if (!($c2r2013Installed)) {
                                $ActionFile = "$scriptPath\$15MSIVBS"
                                if (Test-Path -Path $ActionFile) {
                                    if($ProductsToRemove -ne 'AllOfficeProducts'){
                                        foreach($product in $ProductsToRemove){
                                            switch($product){
                                                "MainOfficeProduct" {
                                                    $ProductName = (Get-OfficeVersion).DisplayName | select -Unique
                                                    if($ProductName -eq $MainOfficeProduct.DisplayName -and $MainOfficeProduct.DisplayName -match '2013'){
                                                        Write-Host "`n`tRemoving "$MainOfficeProduct.DisplayName"..."
                                                        $MainOfficeProductName = $MainOfficeProduct.Name
                                                        $cmdLine = """$ActionFile"" $MainOfficeProductName $argList"
                                                        $cmd = "cmd /c cscript //Nologo $cmdLine"
                                                        Invoke-Expression $cmd
                                                    }
                                                }
                                                "Visio" {
                                                    $ProductName = Get-OfficeVersion -ShowAllInstalledProducts | ? {$_.DisplayName -match $product}
                                                    foreach($prod in $ProductName){
                                                        if($VisioProduct.DisplayName -eq $prod.DisplayName -and $prod.DisplayName -match '2013'){
                                                            Write-Host "`n`tRemoving "$VisioProduct.DisplayName"..."
                                                            $VisioProductName = $VisioProduct.Name
                                                            $cmdLine = """$ActionFile"" $VisioProductName $argList"
                                                            $cmd = "cmd /c cscript //Nologo $cmdLine"
                                                            Invoke-Expression $cmd
                                                        }
                                                    }
                                                }
                                                "Project" {
                                                    $ProductName = Get-OfficeVersion -ShowAllInstalledProducts | ? {$_.DisplayName -match $product}
                                                    foreach($prod in $ProductName){
                                                        if($ProjectProduct.DisplayName -eq $prod.DisplayName -and $prod.DisplayName -match '2013'){
                                                            Write-Host "`n`tRemoving "$ProjectProduct.DisplayName"..."
                                                            $ProjectProductName = $ProjectProduct.Name
                                                            $cmdLine = """$ActionFile"" $ProjectProductName $argList"
                                                            $cmd = "cmd /c cscript //Nologo $cmdLine"
                                                            Invoke-Expression $cmd
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        Write-Host "`n"
                                        $office15Removed = $true
                                    } else {
                                        Write-Host "`n`tRemoving all Office 2013 products..."
                                        $cmdLine = """$ActionFile"" CLIENTALL $argList"
                                        $cmd = "cmd /c cscript //Nologo $cmdLine"
                                        Invoke-Expression $cmd
                                        $office15Removed = $true
                                    }
                                } else {
                                   throw "Required file missing: $ActionFile"
                                }             
                            } else {
                              throw "Office 2013 cannot be removed if 2013 Click-To-Run is installed. Use the -RemoveClickToRunVersions parameter to remove Click-To-Run installs."
                            }
                        }
                    }
                    "16.*"
                    {
                       if (!($office16Removed)) {
                           if ($Remove2016Installs) {
                                if (!($c2r2016Installed)) {
                                    $ActionFile = "$scriptPath\$16MSIVBS"
                                    if (Test-Path -Path $ActionFile) {
                                        if($ProductsToRemove -ne 'AllOfficeProducts'){
                                            foreach($product in $ProductsToRemove){
                                                switch($product){
                                                    "MainOfficeProduct" {
                                                        $ProductName = (Get-OfficeVersion).DisplayName | select -Unique
                                                        if($ProductName -eq $MainOfficeProduct.DisplayName -and $MainOfficeProduct.DisplayName -match '2016'){
                                                            Write-Host "`n`tRemoving "$MainOfficeProduct.DisplayName"..."
                                                            $MainOfficeProductName = $MainOfficeProduct.Name
                                                            $cmdLine = """$ActionFile"" $MainOfficeProductName $argList"
                                                            $cmd = "cmd /c cscript //Nologo $cmdLine"
                                                            Invoke-Expression $cmd
                                                        }
                                                    }
                                                    "Visio" {
                                                        $ProductName = Get-OfficeVersion -ShowAllInstalledProducts | ? {$_.DisplayName -match $product}
                                                        foreach($prod in $ProductName){
                                                            if($VisioProduct.DisplayName -eq $prod.DisplayName -and $prod.DisplayName -match '2016'){
                                                                Write-Host "`n`tRemoving "$VisioProduct.DisplayName"..."
                                                                $VisioProductName = $VisioProduct.Name
                                                                $cmdLine = """$ActionFile"" $VisioProductName $argList"
                                                                $cmd = "cmd /c cscript //Nologo $cmdLine"
                                                                Invoke-Expression $cmd
                                                            }
                                                        }
                                                    }
                                                    "Project" {
                                                        $ProductName = Get-OfficeVersion -ShowAllInstalledProducts | ? {$_.DisplayName -match $product}
                                                        foreach($prod in $ProductName){
                                                            if($ProjectProduct.DisplayName -eq $prod.DisplayName -and $prod.DisplayName -match '2016'){
                                                                Write-Host "`n`tRemoving "$ProjectProduct.DisplayName"..."
                                                                $ProjectProductName = $ProjectProduct.Name
                                                                $cmdLine = """$ActionFile"" $ProjectProductName $argList"
                                                                $cmd = "cmd /c cscript //Nologo $cmdLine"
                                                                Invoke-Expression $cmd
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            Write-Host "`n"
                                            $office16Removed = $true
                                        } else {
                                            Write-Host "`n`tRemoving all Office 2010 products..."
                                            $cmdLine = """$ActionFile"" CLIENTALL $argList"
                                            $cmd = "cmd /c cscript //Nologo $cmdLine"
                                            Invoke-Expression $cmd
                                            $office16Removed = $true
                                        }
                                    } else {
                                       throw "Required file missing: $ActionFile"
                                    }
                                } else {
                                  throw "Office 2016 cannot be removed if 2016 Click-To-Run is installed. Use the -RemoveClickToRunVersions parameter to remove Click-To-Run installs."
                                }
                           }
                       }
                    }
                    default 
                    {
                        continue
                    }
                }
            }
        }

    }
  }
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

    $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet(‘DefaultDisplayPropertySet’,[string[]]$defaultDisplaySet)
    $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
}

process {

 $results = new-object PSObject[] 0;

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

           if ($ConfigItemList.Contains($key.ToUpper()) -and $name.ToUpper().Contains("MICROSOFT OFFICE") `
                                                        -and $name.ToUpper() -notlike "*MUI*" `
                                                        -and $name.ToUpper() -notlike "*VISIO*" `
                                                        -and $name.ToUpper() -notlike "*PROJECT*" `
                                                        -and $name.ToUpper() -notlike "*PROOFING*") {
              $primaryOfficeProduct = $true
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

function GetProductName {
param(
    [Parameter()]
    [string]$ProductName
)
    if($ProductName -eq 'MainOfficeProduct'){
        $ProductName = (Get-OfficeVersion).DisplayName | select -Unique
    } 
        
    $HKLM = [UInt32] "0x80000002"
    $HKCR = [UInt32] "0x80000000"
 
    $installKeys = 'SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall',
                   'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
                   

    $regProv = Get-WmiObject -list "StdRegProv" -namespace root\default -ComputerName $env:COMPUTERNAME

    if($ProductName.ToLower() -match "visio" -or $ProductName.ToLower() -match "project"){
        $ProductName = " " + $ProductName + " "
    }

    foreach ($regKey in $installKeys) {
        $keyList = new-object System.Collections.ArrayList
        $keys = $regProv.EnumKey($HKLM, $regKey)

        foreach ($key in $keys.sNames) {
            $path = Join-Path $regKey $key
            $name = $regProv.GetStringValue($HKLM, $path, "DisplayName").sValue
            $version = $regProv.GetStringValue($HKLM, $path, "DisplayVersion").sValue
            
            if($name){
                if($name.ToLower() -match $ProductName.ToLower()){
                    if($path -notmatch "{.{8}-.{4}-.{4}-.{4}-0000000FF1CE}"){
                        if($key.Split(".")[1] -ne $null){
                            $prodName = $key.Split(".")[1]
                        } else {
                            $prodName = $key
                        }
                        $prodVersion = $version.Split(".")[0]
                        $DisplayName = $name
                    }
                }
            }

            $Result = New-Object -TypeName PSObject
            Add-Member -InputObject $Result -MemberType NoteProperty -Name "DisplayName" -Value $DisplayName 
            Add-Member -InputObject $Result -MemberType NoteProperty -Name "Name" -Value $prodName
            Add-Member -InputObject $Result -MemberType NoteProperty -Name "Version" -Value $prodVersion
        }
    }

    return $Result

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

$dotSourced = IsDotSourced -InvocationLine $MyInvocation.Line

if (!($dotSourced)) {
   Remove-PreviousOfficeInstalls -RemoveClickToRunVersions $RemoveClickToRunVersions -Remove2016Installs $Remove2016Installs -Force $Force -KeepUserSettings $KeepUserSettings -KeepLync $KeepLync -NoReboot $NoReboot
}