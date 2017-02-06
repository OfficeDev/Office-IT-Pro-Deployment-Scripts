﻿Function Remove-OfficeClickToRun {
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
        [string] $TargetFilePath = $NULL
    )

     Process{
 
        $scriptRoot = GetScriptRoot

        newCTRRemoveXml | Out-File $RemoveCTRXmlPath
    
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

            $osVersion = (Get-WmiObject -Class Win32_OperatingSystem).Version
            [int]$osVersion = $osVersion.Split('.')[0]
            if($osVersion -ge '10') {
                Remove-PinnedOfficeApplications
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
                    #write log
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

Function newCTRRemoveXml {
#Create a xml configuration file to remove all Office CTR products.
@"
<Configuration>
  <Remove All="True">
  </Remove>
  <Display Level="None" AcceptEULA="TRUE" />
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

function Remove-PinnedOfficeApplications { 
    [CmdletBinding()] 
    Param( 
        [Parameter()]
        [string]$Action = "Unpin from taskbar"
    ) 

    $ctr = (Get-OfficeVersion).ClickToRun
    $InstallPath = (Get-OfficeVersion).InstallPath
    $officeVersion = (Get-OfficeVersion).Version.Split('.')[0]

    if($ctr -eq $true) {
        $officeAppPath = $InstallPath + "\root\Office" + $officeVersion
    } else {
        $officeAppPath = $InstallPath + "Office" + $officeVersion
    }

    $officeAppList = @("WINWORD.EXE", "EXCEL.EXE", "POWERPNT.EXE", "ONENOTE.EXE", "MSACCESS.EXE", "MSPUB.EXE", "OUTLOOK.EXE",
                       "lync.exe", "GROOVE.EXE", "WINPROJ.EXE", "VISIO.EXE")

    $osVersion = (Get-WmiObject -Class Win32_OperatingSystem).Version
    [int]$osVersion = $osVersion.Split('.')[0]
    
    foreach($app in $officeAppList){
        if(Test-Path ($officeAppPath + "\$app")){
            switch($Action) {
                "Pin to Start" {
                    if($osVersion -ge '10'){
                        $actionId = '51201'
                    } else { 
                        $actionId = '5381'
                    }
                }
                "Unpin from Start" {
                    if($osVersion -ge '10'){
                        $actionId = '51394'
                    } else { 
                        $actionId = '5382'
                    } 
                }
                "Pin to taskbar" {
                    $actionId = '5386'
                }
                "Unpin from taskbar" {
                    $actionId = '5387'
                }   
            }

            InvokeVerb -FilePath ($officeAppPath + "\$app") -Verb $(GetVerb -VerbId $actionId)
            
        }
    }
}

function Remove-PinnedOfficeAppsForWindows10() {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$OfficeApp,

        [Parameter()]
        [string]$Action = 'Unpin from taskbar'
    )

    switch($OfficeApp){
        "WINWORD" {
            $officeAppName = "Word"  
        }
        "EXCEL" {
            $officeAppName = "Excel"
        }
        "POWERPNT" {
            $officeAppName = "PowerPoint"
        }
        "ONENOTE" {
            $officeAppName = "OneNote"
        }
        "MSACCESS" {
            $officeAppName = "Access"
        }
        "MSPUB" { 
            $officeAppName = "Publisher"
        }
        "OUTLOOK" {
            $officeAppName = "Outlook"
        }
        "lync" {
            $officeAppName = "Skype For Business"
        }
        "GROOVE" {
            $officeAppName = "OneDrive For Business"
        }
        "WINPROJ" {
            $officeAppName = "Project"
        }
        "VISIO" {
            $officeAppName = "Visio"
        }
    }

    ((New-Object -Com Shell.Application).NameSpace('shell:::{4234d49b-0245-4df3-b780-3893943456e1}').Items() | ? {$_.Name -like "$officeAppName*"}).Verbs() | ? {$_.Name.replace('&','') -match $Action} | % {$_.DoIt()}
       
}

function InvokeVerb {
    Param(
    [string]$FilePath,
    [string]$verb
    )

    $verb = $verb.Replace("&","") 
    $path = Split-Path $FilePath 
    $shell = New-Object -ComObject "Shell.Application"  
    $folder = $shell.Namespace($path)    
    $item = $folder.Parsename((Split-Path $FilePath -leaf)) 
    $itemVerb = $item.Verbs() | ? {$_.Name.Replace("&","") -eq $verb} 
    
    $osVersion = (Get-WmiObject -Class Win32_OperatingSystem).Version
    [int]$osVersion = $osVersion.Split('.')[0]
    
    if(($itemVerb -eq $null) -and ($osVersion -ge '10')){ 
        Remove-PinnedOfficeAppsForWindows10 -OfficeApp $item.Name -Action $verb             
    } else { 
        if($itemVerb){
            $itemVerb.DoIt() 
        }
    } 
}

function GetVerb { 
    Param(
        [int]$verbId
    ) 
    
    try { 
        $t = [type]"CosmosKey.Util.MuiHelper" 
    } catch { 
        $def = [Text.StringBuilder]"" 
        [void]$def.AppendLine('[DllImport("user32.dll")]') 
        [void]$def.AppendLine('public static extern int LoadString(IntPtr h,uint id, System.Text.StringBuilder sb,int maxBuffer);') 
        [void]$def.AppendLine('[DllImport("kernel32.dll")]') 
        [void]$def.AppendLine('public static extern IntPtr LoadLibrary(string s);') 
        Add-Type -MemberDefinition $def.ToString() -Name MuiHelper -Namespace CosmosKey.Util             
    } 
    if($global:CosmosKey_Utils_MuiHelper_Shell32 -eq $null){         
        $global:CosmosKey_Utils_MuiHelper_Shell32 = [CosmosKey.Util.MuiHelper]::LoadLibrary("shell32.dll") 
    } 

    $maxVerbLength=255 
    $verbBuilder = New-Object Text.StringBuilder "",$maxVerbLength 
    [void][CosmosKey.Util.MuiHelper]::LoadString($CosmosKey_Utils_MuiHelper_Shell32,$verbId,$verbBuilder,$maxVerbLength) 
    
    return $verbBuilder.ToString() 
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
