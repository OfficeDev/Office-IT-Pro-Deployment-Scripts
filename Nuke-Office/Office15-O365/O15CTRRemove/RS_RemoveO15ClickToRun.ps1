
# Copyright © 2012, Microsoft Corporation. All rights reserved.
#============================================================= 

# Load UtilitiesOffScrub_O15c2r.vbs
. .\utils_SetupEnv.ps1
. .\CL_Office.ps1
#$OfficeVersion=""

Import-LocalizedData -BindingVariable Strings_RS_RemoveO15ClickToRun -FileName RS_RemoveO15ClickToRun

# get the content of a log file, return the log that matches $key
#function Get-Off15RemovalResultFromLog($O15RemovalLogPath, $key){ 
  
   #if(!(test-path ($O15RemovalLogPath))){
    #  return $null
   #}
   
  # $logs= get-content $O15RemovalLogPath
   
   #if($logs -eq $null) { return $null }
   
   #$result = (($logs | where { ($_ -match $key)}).Split())[-1]
  # $result = $logs | where { ($_ -match $key)}

   #return $result
#}

# get the content of a log file, returns the entire log 
#function Get-Off15RemovalResultFromLog2($O15RemovalLogPath){   

   #if(!(test-path ($O15RemovalLogPath))){
      #return $null
   #}
   
   #$log= get-content $O15RemovalLogPath   
     
   #if($log -eq $null) { return $null }

   #return $log
#}
#Function to delete folders after reboot

  
#function ExecuteAndCaptureProgress 
#{
  # $job = start-job -scriptblock {CScript.exe .\OffScrubC2R.vbs}

   #while($job.State -eq "running")

  # {  
      #$jbs = receive-job -job $job -keep
  
      #$strbuilder = New-Object system.Text.StringBuilder
      #foreach($jb in $jbs)
      #{  
        # if(($jb -ne "Microsoft (R) Windows Script Host Version 5.8") -and ($jb -ne "Copyright (C) Microsoft Corporation. All rights reserved." ))
         # {
          #if ($jb){         
           #  Write-DiagProgress -activity $Strings_RS_RemoveO15ClickToRun.ID_NAME_progAc_Off15C2R -status  $jb   
      
            # }
          #}
       #}
      #$str = $strbuilder.ToString()
      
    #}
 #}
 #function to unpin the applications from taskbar
 #function Pin-Taskbar([string]$Item = "",[string]$Action = ""){
    #if($Item -eq ""){
       # Write-Error -Message "You need to specify an item" -ErrorAction Stop
    #}
    #if($Action -eq ""){
       # Write-Error -Message "You need to specify an action: Pin or Unpin" -ErrorAction Stop
    #}
   # if((Get-Item -Path "C:\Program Files\Microsoft Office\root\Office16\ONENOTE.EXE" -ErrorAction SilentlyContinue) -eq $null){
       # Write-Error -Message "$Item not found" -ErrorAction Stop
    #}
   # $Shell = New-Object -ComObject "Shell.Application"
   # $ItemParent = Split-Path -Path "C:\Program Files\Microsoft Office\root\Office16\ONENOTE.EXE" -Parent
    #$ItemLeaf = Split-Path -Path "C:\Program Files\Microsoft Office\root\Office16\ONENOTE.EXE" -Leaf
   # $Folder = $Shell.NameSpace($ItemParent)
   # $ItemObject = $Folder.ParseName($ItemLeaf)
    #$Verbs = $ItemObject.Verbs()
   # switch("Unpin"){
         #"Pin"   {$Verb = $Verbs | Where-Object -Property Name -EQ "Pin to Tas&kbar" 
        # $wshell = New-Object -ComObject Wscript.Shell
         # $wshell.Popup("pin",0,"Done",0x1)
        # }
        #"Unpin" {$Verb = $Verbs | Where-Object -Property Name -EQ "Unpin from Tas&kbar"
        #$wshell = New-Object -ComObject Wscript.Shell
       # $wshell.Popup("Unpin",0,"Done",0x1)
        #$verb>>"C:\txtverb.txt"
       # }
        #default {Write-Error -Message "Invalid action, should be Pin or Unpin" -ErrorAction Stop}
       
      # }
   
    #if($Verb -eq $null){
      #  Write-Error -Message "That action is not currently available on this item" -ErrorAction Stop
    #} else {
       # $Result = $Verb.DoIt()
   # }
#}

 #### function to delete the folders after reboot..#
 function Move-LockedFile 
{ 
    param($path, $destination)
    $path = (Resolve-Path $path).Path 
    $destination = $executionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($destination)
    $MOVEFILE_DELAY_UNTIL_REBOOT = 0x00000004

   $memberDefinition = @’ 
    [DllImport("kernel32.dll", SetLastError=true, CharSet=CharSet.Auto)] 
    public static extern bool MoveFileEx(string lpExistingFileName, string lpNewFileName, 
      int dwFlags); 
‘@

    $type = Add-Type -Name MoveFileUtils -MemberDefinition $memberDefinition -PassThru 
    $type::MoveFileEx($path, $destination, $MOVEFILE_DELAY_UNTIL_REBOOT) 
}

#function to delete temp folder files.....
function Temp-FoldersDelete 
{
 Set-Location "C:\Users"
	
	if(test-path "C:\Users\*\Appdata\Local\Temp\OffScrubC2R")
	{
    Remove-Item ".\*\Appdata\Local\Temp\OffScrubC2R" -recurse -force 
	}
	elseif(test-path "C:\Users\*\Appdata\Local\Temp\OffScrub_O16msi" )
	{
    Remove-Item ".\*\Appdata\Local\Temp\OffScrub_O16msi" -recurse -force
	}
	elseif(test-path  "C:\Users\*\Appdata\Local\Temp\OffScrub_O15msi")
	{
    Remove-Item ".\*\Appdata\Local\Temp\OffScrub_O15msi" -recurse -force
	}
}
##### Function to remove office license. Only works in Admin mode ###### 
     
function Remove-L
{ Param ([string]$Key, [string]$OffClass)

$Computer = "."
$Class = $OffClass #"OfficeSoftwareProtectionProduct"

$Method = "UninstallProductKey"

$ID = $Key


$filter="PartialProductKey = '$ID'"

$MC = get-WMIObject -Class $class -computer $Computer -Namespace "ROOT\CIMV2" -filter $filter

$InParams = $mc.psbase.GetMethodParameters($Method)
$mc.PSBase.InvokeMethod($Method ,$Null)

}

#===========================================================================================
# expand-diagcab
#===========================================================================================
 function expand-diagcab($diagcab,$destinationFolder){

     $i = 0
     $oldFolder = $destinationFolder

     while(Test-Path $destinationFolder){
           $destinationFolder   = $oldFolder + "" + $i
           $i = $i + 1
		
     }
     
     md $destinationFolder -force
     & "expand" $diagcab "-f:*.*" $destinationFolder
}

 #===========================================================================================
 #OfficeBitness
 #===========================================================================================
 function OfficeBitness()
{
    $officearchitecture = "HKLM:\SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration" #For O15 SP1
    $officearchitecture1 = "HKLM:\SOFTWARE\Microsoft\Office\15.0\ClickToRun\propertyBag" #For O15


    if(Test-Path $officearchitecture)
    {
        $Key = Get-Item -path $officearchitecture
	
	    if(($Key.GetValue("platform")) -ne $null)
	    {
            $platform = Get-ItemProperty -Path $officearchitecture -Name Platform |%{$_.Platform}
	
        }
	}
	elseif(Test-Path $officearchitecture1) 
	{
	    $Key1 = Get-Item -path $officearchitecture1
	
	    if(($Key1.GetValue("platform")) -ne $null)
	    {
	        $platform = Get-ItemProperty -Path $officearchitecture1 -Name Platform |%{$_.Platform}
		
	    }
	}

	if(($platform -eq "x86") -or ($platform -eq $null))
    {
        $platform = "32bit"
    }
    elseif($platform -eq "x64")
    {
        $platform = "64bit"
    }
  
	return $platform 
}
#===========================================================================================
#Runcleanospp
#===========================================================================================
function RunAppVCleaner()
{
    $OSarchitecture = $env:PROCESSOR_ARCHITECTURE
    $platform = OfficeBitness
	
   if($platform -eq "64bit")
   {

      [string] $currentpath = (Get-Location -PSProvider FileSystem).ProviderPath
	 
      expand-diagcab ("$currentpath\cleanospp.cab") "$currentpath\x64"
      cd .\x64
      .\cleanospp.exe /M:manifest /PID:"9AC08E99-230B-47e8-9721-4577B7F124EA" /RSP:"SOFTWARE\Microsoft\Office\15.0\ClickToRun\appvMachineRegistryStore" /DD:"$env:ProgramW6432\Microsoft Office 15\Data\MachineData" /x64MP:"AppXManifest64.xml" /x64DDC:"DeploymentConfig.xml" /ABD:"$env:ProgramW6432\Microsoft Office 15\root\client" /PD:"$env:ProgramW6432\Microsoft Office 15"
	   .\cleanospp.exe /M:manifest /PID:"9AC08E99-230B-47e8-9721-4577B7F124EA" /RSP:"SOFTWARE\Microsoft\Office\16.0\ClickToRun\appvMachineRegistryStore" /DD:"$env:ProgramW6432\Microsoft Office\Data\MachineData" /x64MP:"AppXManifest64.xml" /x64DDC:"DeploymentConfig.xml" /ABD:"$env:ProgramW6432\Microsoft Office 15\root\client" /PD:"$env:ProgramW6432\Microsoft Office"

      cd ..
   }
   else
   {
      [string] $currentpath = (Get-Location -PSProvider FileSystem).ProviderPath
      expand-diagcab ("$currentpath\cleanospp.cab") "$currentpath\x86"
      cd .\x86
      if(($OSarchitecture -ne "x86") -and (Test-Path "$env:ProgramW6432"))
      {
         .\cleanospp.exe /M:manifest /PID:"9AC08E99-230B-47e8-9721-4577B7F124EA" /RSP:"SOFTWARE\Microsoft\Office\15.0\ClickToRun\appvMachineRegistryStore" /DD:"$env:ProgramW6432\Microsoft Office 15\Data\MachineData" /x86MP:"AppXManifest32.xml" /x86DDC:"DeploymentConfig.xml" /ABD:"$env:ProgramW6432\Microsoft Office 15\root\client" /PD:"$env:ProgramW6432\Microsoft Office 15"
		 .\cleanospp.exe /M:manifest /PID:"9AC08E99-230B-47e8-9721-4577B7F124EA" /RSP:"SOFTWARE\Microsoft\Office\16.0\ClickToRun\appvMachineRegistryStore" /DD:"$env:ProgramW6432\Microsoft Office\Data\MachineData" /x86MP:"AppXManifest32.xml" /x86DDC:"DeploymentConfig.xml" /ABD:"$env:ProgramW6432\Microsoft Office\root\client" /PD:"$env:ProgramW6432\Microsoft Office"

      }
      else
      {
         .\cleanospp.exe /M:manifest /PID:"9AC08E99-230B-47e8-9721-4577B7F124EA" /RSP:"SOFTWARE\Microsoft\Office\15.0\ClickToRun\appvMachineRegistryStore" /DD:"$env:ProgramFiles\Microsoft Office 15\Data\MachineData" /x86MP:"AppXManifest32.xml" /x86DDC:"DeploymentConfig.xml" /ABD:"$env:ProgramFiles\Microsoft Office 15\root\client" /PD:"$env:ProgramFiles\Microsoft Office 15"
		 .\cleanospp.exe /M:manifest /PID:"9AC08E99-230B-47e8-9721-4577B7F124EA" /RSP:"SOFTWARE\Microsoft\Office\16.0\ClickToRun\appvMachineRegistryStore" /DD:"$env:ProgramFiles\Microsoft Office\Data\MachineData" /x86MP:"AppXManifest32.xml" /x86DDC:"DeploymentConfig.xml" /ABD:"$env:ProgramFiles\Microsoft Office\root\client" /PD:"$env:ProgramFiles\Microsoft Office"

      }
      cd ..
   }
}

#call RunAppvCleaner 
#RunAppVCleaner

#$ScrubRetValPath = join-path $env:temp "OffScrubC2R\ScrubRetValFile.txt"
#$ScrubRetVal = Get-Off15RemovalResultFromLog2 $ScrubRetValPath "Removal result"
      
####### MAIN ###########

#if ((Test-Path "HKLM:\SOFTWARE\Microsoft\Office\15.0\ClickToRun") -or (Test-Path "$Env:ProgramFiles\Microsoft Office 15"))
#{
#   Write-DiagProgress -activity $Strings_RS_RemoveO15ClickToRun.ID_Name_RC_Status 
   
   ########### Remove Licenses ###########
   $OfficeAppId = "0ff1ce15-a989-479d-af46-f275c6370663" #Office 2013
   $OS = (Get-WmiObject Win32_OperatingSystem).Name

   $Win =$OS.Substring(0,19)

   if ($Win -eq "Microsoft Windows 7")
   {
   $query = "Select Name, PartialProductKey from OfficeSoftwareProtectionProduct Where ApplicationId = ""$OfficeAppId"" and PartialProductKey <> "" NULL "" "
  
   $Class = "OfficeSoftwareProtectionProduct"
   }else
      {
      $query = "Select Name, PartialProductKey from SoftwareLicensingProduct Where ApplicationId = ""$OfficeAppId"" and PartialProductKey <> "" NULL "" " 
      $Class = "SoftwareLicensingProduct"
      }
   $ProductInstances = gwmi -Query $query -ErrorAction SilentlyContinue
 
   Foreach ($instance in $ProductInstances)
   {
         if ($instance.PartialProductKey -ne $null)
         {
		    #$instance.Name
            #$instance.PartialProductKey
            $O15 = $instance.Name.substring(0,9)
		
			#y only office 15
            if ($O15 -eq "Office 15")
            {
            Remove-L $instance.PartialProductKey -OffClass $Class                        
            }
         }

   }

   ######## MAIN ###########

  Write-DiagProgress -activity $Strings_RS_RemoveO15ClickToRun.ID_NAME_progAc_Off15C2R -status $Strings_RS_RemoveO15ClickToRun.ID_NAME_progSt_Off15C2R

 

  #check all folders path ,registry path of office pack...
  $programfiles="C:\Program Files\Microsoft Office"
  $programfilex86="C:\Program Files (x86)\Microsoft Office"
  $programfiles15="C:\Program Files\Microsoft Office 15"
  $programfilex8615="C:\Program Files (x86)\Microsoft Office 15"
  $ProgramData="C:\ProgramData\Microsoft\Office"
  $commonfiles="C:\Program Files (x86)\Common Files\Microsoft Shared\OFFICE15"
  $commonfiles16="C:\Program Files (x86)\Common Files\Microsoft Shared\OFFICE16"
  $commonfilespr="C:\Program Files\Common Files\Microsoft Shared\OFFICE15"
  $commonfilespr16="C:\Program Files\Common Files\Microsoft Shared\OFFICE16"
  $programdatac2R="C:\ProgramData\Microsoft\ClickToRun"
  $sourceenginefilecheck="C:\Program Files (x86)\Common Files\Microsoft Shared\Source Engine"
  $sourceenginefilecheckpr="C:\Program Files\Common Files\Microsoft Shared\Source Engine"
  $O15C2R= Test-Path  "HKLM:\Software\Microsoft\Office\15.0\ClickToRun"
  $O16C2R= Test-Path  "HKLM:\Software\Microsoft\Office\16.0\ClickToRun"
  $O16C2R10= Test-Path  "HKLM:\Software\Microsoft\Office\ClickToRun"
  $OfficeReg=Test-Path  "HKLM:\Software\Microsoft\Office"
  $OfficeRegHKCU=test-path "HKCU:\Software\Microsoft\Office"
  $officeservies=test-path "HKLM:\SYSTEM\CurrentControlSet\Services\ose"   
  $officeservies64=test-path "HKLM:\SYSTEM\CurrentControlSet\Services\ose64"   
  $offcieWow6432=test-path "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Office"



  #get registry path to delete ClicktoRun EXtensibility component...
  #$regkey15="HKCR:\Installer\Products\00005109C80000000000000000F01FEC";
  # $regkey16="HKCR:\Installer\Products\00006101C80000000000000000F01FEC";
   #$regkey151="HKCR:\Installer\Products\00005109C80000000100000000F01FEC";
  # $regkey161="HKCR:\Installer\Products\00006101C80000000100000000F01FEC";


  #Check whether the offcie istalled is MSI 15,16 or C2R...
  $MSI15 = isOfficeMSI_Installed "Office15"
  $MSI16 = isOfficeMSI_Installed "Office16"
    
  #get the registry path for context menu(NEW) shortcut...
  $regOffWordShortcut="HKCR:\.docx\Word.Document.12\shellNew";
  $regOfficedatabase ="HKCR:\.mdb\Access.MDBFile\shellNew";
  $regOfficedatabase1="HKCR:\.mdb\shellNew";
  $regOffpubdoc ="HKCR:\.pub\publisher.document.15\shellNew";
  $regOffpubdoc1 ="HKCR:\.pub\publisher.document.16\shellNew";	
  $regOfficeXcelShortcut ="HKCR:\.xlsx\Excel.Sheet.12\shellNew";
  $regOffpowerpoint="HKCR:\.pptx\Powerpoint.Show.12\shellNew";
  $regOffmsproject="HKCR:\.mpp\MSProject.Project.9\shellNew";
  $regOffrtf="HKCR:\.rtf\shellNew";
  
  $onenote="C:\Program Files\Microsoft Office\root\Office16\ONENOTE.EXE"
  $excel="C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
  $outlook="C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE" 
  $powerpoint="C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE"
  $msword="C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"

 
  #if c2r installed is true call OffScrubc2R.vbs files to unistall Office C2R files.
  if($O15C2R -or $O16C2R -or $O16C2R10 )
  {
  #calling VBscript
  CScript.exe ".\OffScrubc2r.vbs" //B 
  $ScrubLogDir = join-path $env:temp "OffScrubC2R"

    if(test-path $ScrubLogDir) 
   {   
      $scrubLogFile = (dir $ScrubLogDir | where {$_ -match "_scrubLog.txt"} | sort -prop lastWriteTime | select -last 1).Name

      if ($scrubLogFile -ne $null)
	   {
	           $scrubLogFilePath = join-path $ScrubLogDir $scrubLogFile
       }
   }

   if (test-path $scrubLogFilePath)
   {
      #$RemovalResult = Get-Off15RemovalResultFromLog $scrubLogFilePath "Removal Result"  
	  $RemovalResult= "Success"
   #   if($RemovalResult)  {
   #   #$RemovalResult = $RemovalResult.Split("-")[-1]
   #   $RemovalResult = ($RemovalResult.substring($RemovalResult.indexof('-'))).remove(0,1)
   #   }    
   }
  }

  #if installed office is office MSI 15 then call offscrub_o15msi.vbs 
   elseif($MSI15)
  {
  #$OfficeVersion="15.0"
  $default="CLIENTALL"
  CScript.exe ".\OffScrub_O15msi.vbs" $default 
   #offscrubc2r folder created

     $ScrubLogDir = join-path $env:temp "OffScrub_O15msi"
    if(test-path $ScrubLogDir) 
   {   
      $scrubLogFile = (dir $ScrubLogDir | where {$_ -match "_scrubLog.txt"} | sort -prop lastWriteTime | select -last 1).Name
      if ($scrubLogFile -ne $null)
	   {
	           $scrubLogFilePath = join-path $ScrubLogDir $scrubLogFile
       }
   }
   if (test-path $scrubLogFilePath)
   {
     # $RemovalResult = Get-Off15RemovalResultFromLog $scrubLogFilePath ""  
	 $RemovalResult="Success"  
   }
  }

   #if installed office is office MSI 16 then call offscrub_o16msi.vbs 
  elseif($MSI16)
  {
    #$OfficeVersion="16.0"
    $default="CLIENTALL"
   CScript.exe ".\OffScrub_O16msi.vbs" $default 
   #offscrubc2r folder created

   $ScrubLogDir = join-path $env:temp "OffScrub_O16msi"

    if(test-path $ScrubLogDir) 
   {   
      $scrubLogFile = (dir $ScrubLogDir | where {$_ -match "_scrubLog.txt"} | sort -prop lastWriteTime | select -last 1).Name

      if ($scrubLogFile -ne $null)
	   {
	       $scrubLogFilePath = join-path $ScrubLogDir $scrubLogFile 
		   	
       }
   }
   if (test-path $scrubLogFilePath)
   {
     #$RemovalResult = Get-Off15RemovalResultFromLog $scrubLogFilePath ""  
	 $RemovalResult="Success"   
   }
  }
    

	#if($O15C2R)
	#{
	# $OfficeVersion="15.0"
	#}
	#if($O16C2R)
	#{
	# $OfficeVersion="16.0"
	#}
	#$item= "hklm:\Software\Microsoft\Office\ClickToRun\PropertyBag\"
	#if(test-path $item)
   # {
     #$item=get-itemproperty -path "hklm:\Software\Microsoft\Office\ClickToRun\PropertyBag\" -Name version
    # $stringvalue=$item.Version
     #$16c2R = $stringvalue.Contains("16")
    # $15c2R= $stringvalue.Contains("15")
   # }
	#if($16c2R)
	#{
	#$OfficeVersion="16.0"
	#}
	#elseif($15c2R)
	#{
	# $OfficeVersion="15.0"
	#}

  if(test-path $commonfiles)
  {
  Move-LockedFile "C:\Program Files (x86)\Common Files\Microsoft Shared\OFFICE15" "$null" -ErrorAction SilentlyContinue 
  }
  if(test-path $commonfilespr)
  {
   Move-LockedFile "C:\Program Files\Common Files\Microsoft Shared\OFFICE15" "$null" -ErrorAction SilentlyContinue 
  }
  if(test-path $commonfilespr16)
  {
   Move-LockedFile "C:\Program Files\Common Files\Microsoft Shared\OFFICE16" "$null" -ErrorAction SilentlyContinue 
  }
   if(test-path $programdatac2R)
  {
  Remove-Item -Path "C:\ProgramData\Microsoft\ClickToRun" -force -recurse -ErrorAction SilentlyContinue 
  }
   if(test-path $commonfiles16)
  {
    Move-LockedFile "C:\Program Files (x86)\Common Files\Microsoft Shared\OFFICE16" "$null" -ErrorAction SilentlyContinue 
  }
  if(test-path $sourceenginefilecheck)
  {
  Remove-Item -Path "C:\Program Files (x86)\Common Files\Microsoft Shared\Source Engine" -force -recurse -ErrorAction SilentlyContinue 
  }
   if(test-path $sourceenginefilecheckpr)
  {
  Remove-Item -Path "C:\Program Files\Common Files\Microsoft Shared\Source Engine" -force -recurse -ErrorAction SilentlyContinue 
  }
  
  if(test-path $programfiles)
  {
   # Move-LockedFile "C:\Program Files\Microsoft Office" "$null" -ErrorAction SilentlyContinue 
    Get-ChildItem "C:\Program Files\Microsoft Office" | ? { $_.PSIsContainer } | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue 
    Remove-item -path "C:\Program Files\Microsoft Office" -force -recurse -ErrorAction SilentlyContinue 
   }   
     
  if(test-path $programfilex86)
  {
  #Move-LockedFile "C:\Program Files (x86)\Microsoft Office" "$null" -ErrorAction SilentlyContinue 
  Get-ChildItem "C:\Program Files (x86)\Microsoft Office" | ? { $_.PSIsContainer } | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue 
  Remove-item -path "C:\Program Files (x86)\Microsoft Office" -force -recurse -ErrorAction SilentlyContinue 
  }
         
  if(test-path $ProgramData)
  {
 Remove-item -path "C:\ProgramData\Microsoft\Office" -force -recurse -ErrorAction SilentlyContinue 
  }
  if(test-path $programfiles15)
  {
   # Move-LockedFile "C:\Program Files\Microsoft Office 15" "$null" -ErrorAction SilentlyContinue 
   Get-ChildItem "C:\Program Files\Microsoft Office 15" | ? { $_.PSIsContainer } | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue 
   Remove-item -path "C:\Program Files\Microsoft Office 15" -force -recurse -ErrorAction SilentlyContinue 

  }
 
  if(test-path $programfilex8615)
  {
    #Move-LockedFile "C:\Program Files (x86)\Microsoft Office 15" "$null" -ErrorAction SilentlyContinue 
  Get-ChildItem "C:\Program Files (x86)\Microsoft Office 15"  -Exclude | ? { $_.PSIsContainer } | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue 
   Remove-item -path "C:\Program Files (x86)\Microsoft Office 15" -force -recurse -ErrorAction SilentlyContinue 

  } 

  if(test-path $regOffWordShortcut)
  {
 	Remove-Item -Path HKCR:\.docx\Word.Document.12\shellNew\ -force -recurse -ErrorAction SilentlyContinue 
  }
 
  if(test-path $regOffrtf)
  {
 	Remove-Item -Path HKCR:\.rtf\shellNew\ -force -recurse -ErrorAction SilentlyContinue 
  }
  if(test-path $regOfficedatabase)
  {
 	Remove-Item -Path HKCR:\.mdb\Access.MDBFile\shellNew\ -force -recurse -ErrorAction SilentlyContinue 
  }

   if(test-path $regOfficedatabase1)
  {
 	Remove-Item -Path HKCR:\.mdb\shellNew\ -force -recurse -ErrorAction SilentlyContinue 
  }

   if(test-path $regOffpubdoc)
  {
 	Remove-Item -Path HKCR:\.pub\publisher.document.15\shellNew\ -force -recurse -ErrorAction SilentlyContinue 
  }

   if(test-path $regOffpubdoc1)
  {
 	Remove-Item -Path HKCR:\.pub\publisher.document.16\shellNew\ -force -recurse -ErrorAction SilentlyContinue 
  }

   if(test-path $regOfficeXcelShortcut)
  {
 	Remove-Item -Path HKCR:\.xlsx\Excel.Sheet.12\shellNew\ -force -recurse -ErrorAction SilentlyContinue 
  }

  if(test-path $regOffmsproject)
  {
 	Remove-Item -Path HKCR:\.mpp\MSProject.Project.9\shellNew\ -force -recurse -ErrorAction SilentlyContinue 
  }

  if(test-path $regOffpowerpoint)
  {
 	Remove-Item -Path HKCR:\.pptx\Powerpoint.Show.12\shellNew\ -force -recurse -ErrorAction SilentlyContinue 
  }
   #creating new drivename for HKEY_CLASSES_ROOT...

 New-PSDrive -Name HKCR -PSProvider Registry -Root HKEY_CLASSES_ROOT
 $16regkeydelete= Get-ChildItem -Path HKCR:\Installer\Products\ -Recurse -Include *0000610* 
 $15regkeydelete= Get-ChildItem -Path HKCR:\Installer\Products\ -Recurse -Include *0000510* 
 foreach($del16 in $16regkeydelete)
 {
    $delpath16= $del16.Name
    Remove-Item -Path registry::\$delpath16 -force -recurse 
 } 
  foreach($del15 in $15regkeydelete)
 {
    $delpath15= $del15.Name
    Remove-Item -Path registry::\$delpath15 -force -recurse 
 } 

  if($officeservies)
  {
  Remove-Item -Path HKLM:\SYSTEM\CurrentControlSet\Services\ose -force -recurse -ErrorAction SilentlyContinue 
  }
   if($officeservies64)
  {
    Remove-Item -Path HKLM:\SYSTEM\CurrentControlSet\Services\ose64 -force -recurse -ErrorAction SilentlyContinue 
  }
  if($offcieWow6432)
  {
  Get-ChildItem "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Office"  -Exclude "Common" | ? { $_.PSIsContainer } | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue 
  }
 if($OfficeReg)
{
#Move-LockedFile "HKLM:\Software\Microsoft\Office" "$null" -ErrorAction SilentlyContinue 
Get-ChildItem "HKLM:\Software\Microsoft\Office"  -Exclude "Common" | ? { $_.PSIsContainer } | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue 
Remove-Item -Path HKLM:\Software\Microsoft\Office -force -recurse -ErrorAction SilentlyContinue 
}

if($OfficeRegHKCU)
{
#Move-LockedFile "HKCU:\Software\Microsoft\Office" "$null" -ErrorAction SilentlyContinue 
Get-ChildItem "HKCU:\Software\Microsoft\Office"  -Exclude "Common" | ? { $_.PSIsContainer } | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue 
Remove-Item -Path HKCU:\Software\Microsoft\Office -force -recurse -ErrorAction SilentlyContinue 
}

$OfficeRegHKCU16= test-path "HKCU:\Software\Microsoft\Office\16.0"

if($OfficeRegHKCU16)
{
#Move-LockedFile "HKCU:\Software\Microsoft\Office\16.0" "$null" -ErrorAction SilentlyContinue 
Remove-Item -Path HKCU:\Software\Microsoft\Office\16.0 -force -recurse -ErrorAction SilentlyContinue 
Get-ChildItem "HKCU:\Software\Microsoft\Office\16.0" | ? { $_.PSIsContainer } | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue 
}

$a = Get-WmiObject win32_useraccount -Filter "name = '$env:username'" 
$SID=$a.sid
$OfficeRegHKeyUsers= test-path "registry::\HKEY_USERS\$SID\Software\Microsoft\Office"
#$OfficeRegHKeyUsers16= test-path "registry::\HKEY_USERS\$SID\Software\Microsoft\Office\16.0"

if($OfficeRegHKeyUsers)
{
#Move-LockedFile "registry::\HKEY_USERS\$SID\Software\Microsoft\Office" "$null" -ErrorAction SilentlyContinue 
Remove-Item -Path registry::\HKEY_USERS\$SID\Software\Microsoft\Office -force -recurse -ErrorAction SilentlyContinue 
Get-ChildItem "registry::\HKEY_USERS\$SID\Software\Microsoft\Office" -Exclude "Common" | ? { $_.PSIsContainer } | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue 
}

#if($OfficeRegHKeyUsers16)
#{
#Move-LockedFile "registry::\HKEY_USERS\$SID\Software\Microsoft\Office\16.0" "$null" -ErrorAction SilentlyContinue 
#Remove-Item -Path registry::\HKEY_USERS\$SID\Software\Microsoft\Office\16.0 -force -recurse -ErrorAction SilentlyContinue 
#Get-ChildItem "registry::\HKEY_USERS\$SID\Software\Microsoft\Office\16.0" | ? { $_.PSIsContainer } | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue 
#}

$username=$env:username
$check="C:\Users\$username\AppData\Local\Microsoft\Office"
#$check16="C:\Users\$username\AppData\Local\Microsoft\\Office\16.0"
$Localpath= test-path $check
#$Localpath16= test-path $check16

if($Localpath)
{
#Move-LockedFile $check "$null" -ErrorAction SilentlyContinue 
Remove-Item -Path $check -force -recurse -ErrorAction SilentlyContinue 
Get-ChildItem $check -Exclude "Common" | ? { $_.PSIsContainer } | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue 
}
#if($Localpath16)
#{
#Move-LockedFile $check16 "$null" -ErrorAction SilentlyContinue 
#Remove-Item -Path $check16 -force -recurse -ErrorAction SilentlyContinue 
#Get-ChildItem $check16 -Exclude "Common" | ? { $_.PSIsContainer } | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue 
#}

  $O15reg= Test-Path  "HKLM:\Software\Microsoft\Office\15.0"
  $O16reg= Test-Path  "HKLM:\Software\Microsoft\Office\16.0"


if($O15reg)
{
 Move-LockedFile "HKLM:\Software\Microsoft\Office\15.0" "$null" -ErrorAction SilentlyContinue 
 #Remove-Item -Path HKLM:\Software\Microsoft\Office\15.0 -force -recurse
}
if($O16reg)
{
 Move-LockedFile "HKLM:\Software\Microsoft\Office\16.0" "$null" -ErrorAction SilentlyContinue 
#Remove-Item -Path HKLM:\Software\Microsoft\Office\16.0 -force -recurse
}

  $MSI15 = isOfficeMSI_Installed "Office15"
  $MSI16 = isOfficeMSI_Installed "Office16"

  #get all the empty office folders and delete that. After successfull deletion display successfull and  reboot message to the user..
  if(!(test-path $O15C2R) -and !(test-path $O16C2R) -and !(test-path $O16C2R10) -and !($MSI15) -and !($MSI16))
  {
    #get-DiagInput -id "INT_O15C2R_Uninstalled"
	get-DiagInput -id "INT_O15C2R_Reboot"           
    $RemovalResult| convertto-xml | update-diagreport  -id "RS_RemoveO15ClickToRun" -name "Success:" -Verbosity Informational
	#remove all temp folder files...
	Temp-FoldersDelete    
	$Global:FixedC2R = $true
  }
  #test the path of the office folders and registry paths. If it is present then diplay a error message.
  elseif((test-path $programfiles) -or (test-path $programfilex86) -or (test-path $ProgramData) -or (test-path $O15C2R) -or (test-path $O16C2R) -or (test-path $O16C2R10) -or ($MSI15) -or ($MSI16))
  {
    get-DiagInput -id "INT_O15C2R_UninstallFailed" -Parameter @{'failureMsg' =  "Failed"}        
    "Failed"| convertto-xml | update-diagreport  -id "RS_RemoveO15ClickToRun" -name "Failed:" -Verbosity Informational
	Temp-FoldersDelete 
    $Global:FixedC2R =$false
  }
 
	
   ## Based on the removal result determine the INT/Msg to show    
  # if ($RemovalResult.substring($RemovalResult.Length - 7,7) -ieq "SUCCESS") 
     # {
	 
      #get-DiagInput -id "INT_O15C2R_Uninstalled"
     # $Global:FixedC2R = $true
	   #} 
	   #else
      #{

        # if ($RemovalResult.substring($RemovalResult.Length - 47,47) -ieq "Uninstall requires a system reboot to complete.")
         #{
            #C2R uninstall success - but requires a reboot.
            #get-DiagInput -id "INT_O15C2R_Reboot"          
           # $RemovalResult| convertto-xml | update-diagreport  -id "RS_RemoveO15ClickToRun" -name "Removal result:" -Verbosity Informational
           # $Global:FixedC2R = $true
		#}
         # else
         #{
	        #get-DiagInput -id "INT_O15C2R_UninstallFailed" -Parameter @{'failureMsg' =  $RemovalResult}        
           # $RemovalResult| convertto-xml | update-diagreport  -id "RS_RemoveO15ClickToRun" -name "Removal result:" -Verbosity Informational

        # }

		 #C2R uninstall success - but requires a reboot.
           # get-DiagInput -id "INT_O15C2R_Reboot"          
          # $RemovalResult| convertto-xml | update-diagreport  -id "RS_RemoveO15ClickToRun" -name "Removal result:" -Verbosity Informational
          # $Global:FixedC2R = $true
   #}
 
 
#}  
 
 




# SIG # Begin signature block
# MIIa5AYJKoZIhvcNAQcCoIIa1TCCGtECAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUCjZyykkHqiy+RyfDrC2m8R5U
# HvygghWCMIIEwzCCA6ugAwIBAgITMwAAAHD0GL8jIfxQnQAAAAAAcDANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTUwMzIwMTczMjAy
# WhcNMTYwNjIwMTczMjAyWjCBszELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjENMAsGA1UECxMETU9QUjEnMCUGA1UECxMebkNpcGhlciBEU0UgRVNO
# OkY1MjgtMzc3Ny04QTc2MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAoxTZ7xygeRG9
# LZoEnSM0gqVCHSsA0dIbMSnIKivzLfRui93iG/gT9MBfcFOv5zMPdEoHFGzcKAO4
# Kgp4xG4gjguAb1Z7k/RxT8LTq8bsLa6V0GNnsGSmNAMM44quKFICmTX5PGTbKzJ3
# wjTuUh5flwZ0CX/wovfVkercYttThkdujAFb4iV7ePw9coMie1mToq+TyRgu5/YK
# VA6YDWUGV3eTka+Ur4S+uG+thPT7FeKT4thINnVZMgENcXYAlUlpbNTGNjpaMNDA
# ynOJ5pT2Ix4SYFEACMHe2j9IhO21r9TTmjiVqbqjWLV4aEa/D4xjcb46Q0NZEPBK
# unvW5QYT3QIDAQABo4IBCTCCAQUwHQYDVR0OBBYEFG3P87iErvfMdr24e6w9l2GB
# dCsnMB8GA1UdIwQYMBaAFCM0+NlSRnAK7UD7dvuzK7DDNbMPMFQGA1UdHwRNMEsw
# SaBHoEWGQ2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3Rz
# L01pY3Jvc29mdFRpbWVTdGFtcFBDQS5jcmwwWAYIKwYBBQUHAQEETDBKMEgGCCsG
# AQUFBzAChjxodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY3Jv
# c29mdFRpbWVTdGFtcFBDQS5jcnQwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZI
# hvcNAQEFBQADggEBAF46KvVn9AUwKt7hue9n/Cr/bnIpn558xxPDo+WOPATpJhVN
# 98JnglwKW8UK7lXwoy2Ooh2isywt0BHimioB0TAmZ6GmbokxHG7dxHFU8Ami3cHW
# NnPADP9VCGv8oZT9XSwnIezRIwbcBCzvuQLbA7tHcxgK632ZzV8G4Ij3ipPFEhEb
# 81KVo3Kg0ljZwyzia3931GNT6oK4L0dkKJjHgzvxayhh+AqIgkVSkumDJklct848
# mn+voFGTxby6y9ErtbuQGQqmp2p++P0VfkZEh6UG1PxKcDjG6LVK9NuuL+xDyYmi
# KMVV2cG6W6pgu6W7+dUCjg4PbcI1cMCo7A2hsrgwggTsMIID1KADAgECAhMzAAAB
# Cix5rtd5e6asAAEAAAEKMA0GCSqGSIb3DQEBBQUAMHkxCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xIzAhBgNVBAMTGk1pY3Jvc29mdCBDb2RlIFNp
# Z25pbmcgUENBMB4XDTE1MDYwNDE3NDI0NVoXDTE2MDkwNDE3NDI0NVowgYMxCzAJ
# BgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25k
# MR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xDTALBgNVBAsTBE1PUFIx
# HjAcBgNVBAMTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjCCASIwDQYJKoZIhvcNAQEB
# BQADggEPADCCAQoCggEBAJL8bza74QO5KNZG0aJhuqVG+2MWPi75R9LH7O3HmbEm
# UXW92swPBhQRpGwZnsBfTVSJ5E1Q2I3NoWGldxOaHKftDXT3p1Z56Cj3U9KxemPg
# 9ZSXt+zZR/hsPfMliLO8CsUEp458hUh2HGFGqhnEemKLwcI1qvtYb8VjC5NJMIEb
# e99/fE+0R21feByvtveWE1LvudFNOeVz3khOPBSqlw05zItR4VzRO/COZ+owYKlN
# Wp1DvdsjusAP10sQnZxN8FGihKrknKc91qPvChhIqPqxTqWYDku/8BTzAMiwSNZb
# /jjXiREtBbpDAk8iAJYlrX01boRoqyAYOCj+HKIQsaUCAwEAAaOCAWAwggFcMBMG
# A1UdJQQMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBSJ/gox6ibN5m3HkZG5lIyiGGE3
# NDBRBgNVHREESjBIpEYwRDENMAsGA1UECxMETU9QUjEzMDEGA1UEBRMqMzE1OTUr
# MDQwNzkzNTAtMTZmYS00YzYwLWI2YmYtOWQyYjFjZDA1OTg0MB8GA1UdIwQYMBaA
# FMsR6MrStBZYAck3LjMWFrlMmgofMFYGA1UdHwRPME0wS6BJoEeGRWh0dHA6Ly9j
# cmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3RzL01pY0NvZFNpZ1BDQV8w
# OC0zMS0yMDEwLmNybDBaBggrBgEFBQcBAQROMEwwSgYIKwYBBQUHMAKGPmh0dHA6
# Ly93d3cubWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljQ29kU2lnUENBXzA4LTMx
# LTIwMTAuY3J0MA0GCSqGSIb3DQEBBQUAA4IBAQCmqFOR3zsB/mFdBlrrZvAM2PfZ
# hNMAUQ4Q0aTRFyjnjDM4K9hDxgOLdeszkvSp4mf9AtulHU5DRV0bSePgTxbwfo/w
# iBHKgq2k+6apX/WXYMh7xL98m2ntH4LB8c2OeEti9dcNHNdTEtaWUu81vRmOoECT
# oQqlLRacwkZ0COvb9NilSTZUEhFVA7N7FvtH/vto/MBFXOI/Enkzou+Cxd5AGQfu
# FcUKm1kFQanQl56BngNb/ErjGi4FrFBHL4z6edgeIPgF+ylrGBT6cgS3C6eaZOwR
# XU9FSY0pGi370LYJU180lOAWxLnqczXoV+/h6xbDGMcGszvPYYTitkSJlKOGMIIF
# vDCCA6SgAwIBAgIKYTMmGgAAAAAAMTANBgkqhkiG9w0BAQUFADBfMRMwEQYKCZIm
# iZPyLGQBGRYDY29tMRkwFwYKCZImiZPyLGQBGRYJbWljcm9zb2Z0MS0wKwYDVQQD
# EyRNaWNyb3NvZnQgUm9vdCBDZXJ0aWZpY2F0ZSBBdXRob3JpdHkwHhcNMTAwODMx
# MjIxOTMyWhcNMjAwODMxMjIyOTMyWjB5MQswCQYDVQQGEwJVUzETMBEGA1UECBMK
# V2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0
# IENvcnBvcmF0aW9uMSMwIQYDVQQDExpNaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBD
# QTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBALJyWVwZMGS/HZpgICBC
# mXZTbD4b1m/My/Hqa/6XFhDg3zp0gxq3L6Ay7P/ewkJOI9VyANs1VwqJyq4gSfTw
# aKxNS42lvXlLcZtHB9r9Jd+ddYjPqnNEf9eB2/O98jakyVxF3K+tPeAoaJcap6Vy
# c1bxF5Tk/TWUcqDWdl8ed0WDhTgW0HNbBbpnUo2lsmkv2hkL/pJ0KeJ2L1TdFDBZ
# +NKNYv3LyV9GMVC5JxPkQDDPcikQKCLHN049oDI9kM2hOAaFXE5WgigqBTK3S9dP
# Y+fSLWLxRT3nrAgA9kahntFbjCZT6HqqSvJGzzc8OJ60d1ylF56NyxGPVjzBrAlf
# A9MCAwEAAaOCAV4wggFaMA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFMsR6MrS
# tBZYAck3LjMWFrlMmgofMAsGA1UdDwQEAwIBhjASBgkrBgEEAYI3FQEEBQIDAQAB
# MCMGCSsGAQQBgjcVAgQWBBT90TFO0yaKleGYYDuoMW+mPLzYLTAZBgkrBgEEAYI3
# FAIEDB4KAFMAdQBiAEMAQTAfBgNVHSMEGDAWgBQOrIJgQFYnl+UlE/wq4QpTlVnk
# pDBQBgNVHR8ESTBHMEWgQ6BBhj9odHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtp
# L2NybC9wcm9kdWN0cy9taWNyb3NvZnRyb290Y2VydC5jcmwwVAYIKwYBBQUHAQEE
# SDBGMEQGCCsGAQUFBzAChjhodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2Nl
# cnRzL01pY3Jvc29mdFJvb3RDZXJ0LmNydDANBgkqhkiG9w0BAQUFAAOCAgEAWTk+
# fyZGr+tvQLEytWrrDi9uqEn361917Uw7LddDrQv+y+ktMaMjzHxQmIAhXaw9L0y6
# oqhWnONwu7i0+Hm1SXL3PupBf8rhDBdpy6WcIC36C1DEVs0t40rSvHDnqA2iA6VW
# 4LiKS1fylUKc8fPv7uOGHzQ8uFaa8FMjhSqkghyT4pQHHfLiTviMocroE6WRTsgb
# 0o9ylSpxbZsa+BzwU9ZnzCL/XB3Nooy9J7J5Y1ZEolHN+emjWFbdmwJFRC9f9Nqu
# 1IIybvyklRPk62nnqaIsvsgrEA5ljpnb9aL6EiYJZTiU8XofSrvR4Vbo0HiWGFzJ
# NRZf3ZMdSY4tvq00RBzuEBUaAF3dNVshzpjHCe6FDoxPbQ4TTj18KUicctHzbMrB
# 7HCjV5JXfZSNoBtIA1r3z6NnCnSlNu0tLxfI5nI3EvRvsTxngvlSso0zFmUeDord
# EN5k9G/ORtTTF+l5xAS00/ss3x+KnqwK+xMnQK3k+eGpf0a7B2BHZWBATrBC7E7t
# s3Z52Ao0CW0cgDEf4g5U3eWh++VHEK1kmP9QFi58vwUheuKVQSdpw5OPlcmN2Jsh
# rg1cnPCiroZogwxqLbt2awAdlq3yFnv2FoMkuYjPaqhHMS+a3ONxPdcAfmJH0c6I
# ybgY+g5yjcGjPa8CQGr/aZuW4hCoELQ3UAjWwz0wggYHMIID76ADAgECAgphFmg0
# AAAAAAAcMA0GCSqGSIb3DQEBBQUAMF8xEzARBgoJkiaJk/IsZAEZFgNjb20xGTAX
# BgoJkiaJk/IsZAEZFgltaWNyb3NvZnQxLTArBgNVBAMTJE1pY3Jvc29mdCBSb290
# IENlcnRpZmljYXRlIEF1dGhvcml0eTAeFw0wNzA0MDMxMjUzMDlaFw0yMTA0MDMx
# MzAzMDlaMHcxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
# VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xITAf
# BgNVBAMTGE1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQTCCASIwDQYJKoZIhvcNAQEB
# BQADggEPADCCAQoCggEBAJ+hbLHf20iSKnxrLhnhveLjxZlRI1Ctzt0YTiQP7tGn
# 0UytdDAgEesH1VSVFUmUG0KSrphcMCbaAGvoe73siQcP9w4EmPCJzB/LMySHnfL0
# Zxws/HvniB3q506jocEjU8qN+kXPCdBer9CwQgSi+aZsk2fXKNxGU7CG0OUoRi4n
# rIZPVVIM5AMs+2qQkDBuh/NZMJ36ftaXs+ghl3740hPzCLdTbVK0RZCfSABKR2YR
# JylmqJfk0waBSqL5hKcRRxQJgp+E7VV4/gGaHVAIhQAQMEbtt94jRrvELVSfrx54
# QTF3zJvfO4OToWECtR0Nsfz3m7IBziJLVP/5BcPCIAsCAwEAAaOCAaswggGnMA8G
# A1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFCM0+NlSRnAK7UD7dvuzK7DDNbMPMAsG
# A1UdDwQEAwIBhjAQBgkrBgEEAYI3FQEEAwIBADCBmAYDVR0jBIGQMIGNgBQOrIJg
# QFYnl+UlE/wq4QpTlVnkpKFjpGEwXzETMBEGCgmSJomT8ixkARkWA2NvbTEZMBcG
# CgmSJomT8ixkARkWCW1pY3Jvc29mdDEtMCsGA1UEAxMkTWljcm9zb2Z0IFJvb3Qg
# Q2VydGlmaWNhdGUgQXV0aG9yaXR5ghB5rRahSqClrUxzWPQHEy5lMFAGA1UdHwRJ
# MEcwRaBDoEGGP2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1
# Y3RzL21pY3Jvc29mdHJvb3RjZXJ0LmNybDBUBggrBgEFBQcBAQRIMEYwRAYIKwYB
# BQUHMAKGOGh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljcm9z
# b2Z0Um9vdENlcnQuY3J0MBMGA1UdJQQMMAoGCCsGAQUFBwMIMA0GCSqGSIb3DQEB
# BQUAA4ICAQAQl4rDXANENt3ptK132855UU0BsS50cVttDBOrzr57j7gu1BKijG1i
# uFcCy04gE1CZ3XpA4le7r1iaHOEdAYasu3jyi9DsOwHu4r6PCgXIjUji8FMV3U+r
# kuTnjWrVgMHmlPIGL4UD6ZEqJCJw+/b85HiZLg33B+JwvBhOnY5rCnKVuKE5nGct
# xVEO6mJcPxaYiyA/4gcaMvnMMUp2MT0rcgvI6nA9/4UKE9/CCmGO8Ne4F+tOi3/F
# NSteo7/rvH0LQnvUU3Ih7jDKu3hlXFsBFwoUDtLaFJj1PLlmWLMtL+f5hYbMUVbo
# nXCUbKw5TNT2eb+qGHpiKe+imyk0BncaYsk9Hm0fgvALxyy7z0Oz5fnsfbXjpKh0
# NbhOxXEjEiZ2CzxSjHFaRkMUvLOzsE1nyJ9C/4B5IYCeFTBm6EISXhrIniIh0EPp
# K+m79EjMLNTYMoBMJipIJF9a6lbvpt6Znco6b72BJ3QGEe52Ib+bgsEnVLaxaj2J
# oXZhtG6hE6a/qkfwEm/9ijJssv7fUciMI8lmvZ0dhxJkAj0tr1mPuOQh5bWwymO0
# eFQF1EEuUKyUsKV4q7OglnUa2ZKHE3UiLzKoCG6gW4wlv6DvhMoh1useT8ma7kng
# 9wFlb4kLfchpyOZu6qeXzjEp/w7FW1zYTRuh2Povnj8uVRZryROj/TGCBMwwggTI
# AgEBMIGQMHkxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
# VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xIzAh
# BgNVBAMTGk1pY3Jvc29mdCBDb2RlIFNpZ25pbmcgUENBAhMzAAABCix5rtd5e6as
# AAEAAAEKMAkGBSsOAwIaBQCggeUwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQw
# HAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFJXa
# gk1UKKAsskk1Xtiq37d0YT7EMIGEBgorBgEEAYI3AgEMMXYwdKBagFgATwBmAGYA
# aQBjAGUAVQBuAGkAbgBzAHQAYQBsAGwAZQByAF8AUgBTAF8AUgBlAG0AbwB2AGUA
# TwAxADUAQwBsAGkAYwBrAFQAbwBSAHUAbgAuAHAAcwAxoRaAFGh0dHA6Ly9taWNy
# b3NvZnQuY29tMA0GCSqGSIb3DQEBAQUABIIBAG8DeXUgA88ybqiKfoNVA7hcJwxW
# UzzmTeLBwNDOvkvZA6b5srNzpGHPoXfar2RB1S/HhEuz0aaeGgTPPK9oHkvMxTXz
# P2/VA1oGINlwfsrCnoylnbO7BUt5cS8Db7N6qYB4Cfd2yz1SbfPQGOf9a3OeoNMX
# QOIu843uWJaX5E3TSajr2U2VhMRHuYxqotPUsXIcU9LfokRglrrL2KFCa5yZpT2v
# RIeEX4+dMIvBuWKQRKZsks/AZlWJBtg5vssQggo8eHHOQx8ipO9Kfu2vNM1Hv2Qa
# +wmMOG76OXCQcs077d6RvIGRgfrUy5bPG83B3xxLXRwe0kiCnbtFDk1v8bGhggIo
# MIICJAYJKoZIhvcNAQkGMYICFTCCAhECAQEwgY4wdzELMAkGA1UEBhMCVVMxEzAR
# BgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
# Y3Jvc29mdCBDb3Jwb3JhdGlvbjEhMB8GA1UEAxMYTWljcm9zb2Z0IFRpbWUtU3Rh
# bXAgUENBAhMzAAAAcPQYvyMh/FCdAAAAAABwMAkGBSsOAwIaBQCgXTAYBgkqhkiG
# 9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xNTA3MjAyMTQ1MzBa
# MCMGCSqGSIb3DQEJBDEWBBQjJ02oX3IAQt/7DRJzinV2pSw9oTANBgkqhkiG9w0B
# AQUFAASCAQBrUQeX/UbsfNXAOr+e9Uoz6s9SwZS0xyfpf/viR3a7vouXL+DcvR+h
# IRjgCWJeGJ1+BMDNCXWRX9BA6azo6ZIoACskGJsxG1ZbgvRlZuQ9lMrHc2u5WAeD
# GuqdR/ZObwSI5LiQG2LE4sqBKqZgEd6XktYN2PaP/NUy8STQCUWIWejW6j3fpExI
# QjWwjjkhIToBqp4tSahRTRjzIVqrR8a32IYsvp5HP0Ty7Eo6shqu2lLZmBznQI9m
# z5wBnghcfddeLvW8ji/2UZ7vSQxXytjEAjPG94h6Nvnggw7Ks9oW8x6D51/xOdlb
# amkTtM/+BFmAYEeKblO/vPDjAPOo0+rp
# SIG # End signature block
