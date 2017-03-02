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
        [string[]]$C2RProductsToRemove = "All"
    )

     Process{
 
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
                #write log
                $lineNum = Get-CurrentLineNumber    
                $filName = Get-CurrentFileName 
                WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Please wait while $c2rName is being uninstalled..."
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
                    <# write log#>
                    $lineNum = Get-CurrentLineNumber    
                    $filName = Get-CurrentFileName 
                    WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Office Click-to-Run has been successfully uninstalled." 
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

Function GetScriptRoot() {
 process {
     [string]$scriptPath = "."

     if ($PSScriptRoot) {
       $scriptPath = $PSScriptRoot
     } else {
       $scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
       if (!($scriptPath)) {
          $scriptPath = (Get-Location).Path
       }
     }

     return $scriptPath
 }
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
        $fileName = $_.InvocationInfo.ScriptName.Substring($_.InvocationInfo.ScriptName.LastIndexOf("\")+1)
        WriteToLogFile -LNumber $_.InvocationInfo.ScriptLineNumber -FName $fileName -ActionError $_
    }
}

function Get-CurrentLineNumber {
    $MyInvocation.ScriptLineNumber
}

function Get-CurrentFileName{
    $MyInvocation.ScriptName.Substring($MyInvocation.ScriptName.LastIndexOf("\")+1)
}

function Get-CurrentFunctionName {
    (Get-Variable MyInvocation -Scope 1).Value.MyCommand.Name;
}

Function WriteToLogFile() {
    param( 
      [Parameter(Mandatory=$true)]
      [string]$LNumber,
      [Parameter(Mandatory=$true)]
      [string]$FName,
      [Parameter(Mandatory=$true)]
      [string]$ActionError
    )
    try{
        $headerString = "Time".PadRight(30, ' ') + "Line Number".PadRight(15,' ') + "FileName".PadRight(60,' ') + "Action"
        $stringToWrite = $(Get-Date -Format G).PadRight(30, ' ') + $($LNumber).PadRight(15, ' ') + $($FName).PadRight(60,' ') + $ActionError

        #check if file exists, create if it doesn't
        $getCurrentDatePath = "C:\Windows\Temp\" + (Get-Date -Format u).Substring(0,10)+"OfficeAutoScriptLog.txt"
        if(Test-Path $getCurrentDatePath){#if exists, append 
             Add-Content $getCurrentDatePath $stringToWrite
        }
        else{#if not exists, create new
             Add-Content $getCurrentDatePath $headerString
             Add-Content $getCurrentDatePath $stringToWrite
        }
    } catch [Exception]{
        Write-Host $_
    }
}