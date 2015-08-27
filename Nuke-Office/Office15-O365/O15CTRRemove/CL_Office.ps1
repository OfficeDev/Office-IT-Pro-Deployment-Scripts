# Common Utilities for Microsoft Office
# Get-OfficeAddins: Gives you a list of Office Addins
# Get-OfficePath: Gives you the folder path of the Office installations
#+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

function Get-OfficeAddins {
	Param ([Microsoft.Win32.RegistryHive] $RegKeyHive, [string] $RegKeyPath, [string] $Product)
#write-diagprogress -Activity "Inside Get OfficeAddins..."
$InteractionChoices = ""
$Officeaddins = ""
$type = $RegKeyHive # [Microsoft.Win32.RegistryHive]::CurrentUser 
$regKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($type, $Srv)

$regKey = $regKey.OpenSubKey($RegKeyPath)
if ($regKey.GetSubKeyNames -ne $null)
{
	Foreach($sub in $regKey.GetSubKeyNames())
		{
			$i = $RegKeyPath + $sub
			$regKey1 = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($type, $Srv)
			$regKey1 = $regKey1.OpenSubKey($i)
			Foreach($val in $regKey1.GetValueNames())
			{
			
				if($val	-eq "FriendlyName")
				{
				#$InteractionChoices += System.Environment.NewLine
				$InteractionChoices += [char]13 + [char]10 + $Product + " " + $regKey1.GetValue($val)
			
				}
			}
		}
		
}
	$InteractionChoices
}

#function Get-OfficePath{
##		Param ([string] $Product)
## change code to check for 32bit first then 64bit
#
#$Reg = (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurreontVersion\"App Paths"\Winword.exe)
#
#if ($Reg.Path.EndsWith("OFFICE11\"))
#{	$O_Ver = "Office11"  }
#
#if ($Reg.Path.EndsWith("Office12\"))
#{	$O_Ver = "Office12"  }
#
#if ($Reg.Path.EndsWith("Office14\"))
#{	$O_Ver = "Office14"  }
#
#
#return $O_Ver
#
#}
function Get-OfficePath{
#		Param ([string] $Product)

$TestReg = Test-Path -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\"App Paths"\Winword.exe
if ($TestReg -eq $true)
{
$Reg = (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\"App Paths"\Winword.exe)
}
else 
	{
    if (Test-Path -Path HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\"App Paths"\Winword.exe)
    {$Reg = (Get-ItemProperty HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\"App Paths"\Winword.exe)}
	}

if ($Reg)
{
	switch (($Reg.Path.Split("\"))[-2])
	{
		"OFFICE11" {$O_Ver = "Office11" }
		"Office12" {$O_Ver = "Office12" }
		"Office14" {$O_Ver = "Office14" }
		"Office15" {$O_Ver = "Office15" }
		"Office16" {$O_Ver = "Office16" }
	}
}
 return $O_Ver
}

# Get the version of an office product 
# Usage example: Get-OfficeProductVersion "Outlook.exe"
# Winword.exe, excel.exe, MSAccess.exe, powerpnt.exe, oneNote.exe, infopath.exe, MSPub.exe, visio.exe, . . .

function Get-OfficeProductVersion {
		Param ([string] $Product)

$productPathKey =  "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\$Product"
if (test-path $productPathKey) {
 $productPath = (Get-ItemProperty $productPathKey).path
}
else {
	$productPathKey = "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\App Paths\$Product"
	if (Test-Path $productPathKey){
	 $productPath = (Get-ItemProperty $productPathKey).path
	}
}

if ($productPath) {
	return (Split-Path $productPath -Leaf) 
	}			
}




function Pop-Msg {
	 param([string]$msg ="message",
	 [string]$ttl = "Title",
	 [int]$type = 64) 
	 $popwin = new-object -comobject wscript.shell
	 $null = $popwin.popup($msg,0,$ttl,$type)
	 remove-variable popwin
}


function Off_Exists{
			PARAM ([string] $Product)

$reg = "REGISTRY::HKEY_CLASSES_ROOT\" + $Product +".Application"
# Check normal Office Suites

if ((Test-Path $reg ))
{ 
		#return 1 if product is installed
	   return 1
  
# if not found then check for Office Click 2 Run App-V suites
} Elseif (Test-Path HKLM:\SOFTWARE\Microsoft\SoftGrid)
	{
		#return 1 if product is installed
		return 1
		
	} else { 
	#return 0 if product is installed
	return 0
	}
       
}

# Determine whether office (msi or clickToRun)is installed  or not
# Usage example: Get_OfficeExists "Outlook"
function Get_OfficeExists
{
	PARAM ([string] $Product)
	#Clear-Host
	$reg = "REGISTRY::HKEY_CLASSES_ROOT\" + $Product + ".Application"
	# Check normal Office Suites
	if ((Test-Path $reg ))
	{ 
		#product is installed. return 1
		return 1   
	   
	# if not found then check for Office Click 2 Run App-V suites
	} 
	Elseif ((Test-Path HKLM:\SOFTWARE\Microsoft\SoftGrid) -or (Test-Path HKLM:\SOFTWARE\Wow6432Node\Microsoft\SoftGrid))
	{
		#product is installed. return 1. 
		return 1		
	} 
	else 
	{ 
		#product isn't installed. return 0 
		return 0
	}       
}

#Check if Office product is installed 
# Eg: Check-Office "WinWord.exe" "Word"
#	  Check-Office "Outlook.exe" "Outlook"	  
function Check-Office
{PARAM ([string]$ProductNameWithEXE,[string]$OnlyProductName)
#    if(!$global:__OfficeChecked)
#    {
        $OfficeVersion = @("15.0","14.0","12.0","11.0","16.00")
        $OfficeInstalled = $false
        $OfficeSupported = $false

        foreach ($version in $OfficeVersion)
        {
            $OfficeRegPaths = @("HKLM:\Software\Microsoft\Office\$version","HKLM:\Software\Wow6432Node\Microsoft\Office\$version")
            foreach($OfficeRegPath in $OfficeRegPaths)
            {
                $RealOfficeRegPath = $OfficeRegPath+ "\" + $OnlyProductName + "\InstallRoot"
                if(Test-Path $RealOfficeRegPath)
                {
                    $OfficeInstalled = $true
                    $OfficePath = (Get-item $RealOfficeRegPath).getvalue("Path") + $ProductNameWithEXE
                    $result = CompareVersion -Path $OfficePath -SpecificVersion @(11,0,8169,0)
                    Set-OfficeRegPath -Version $version -OfficeRegPath $OfficeRegPath
                    if($result)
                    {
                        $OfficeSupported = $true
                        break
                    }
                }
            }

            if($OfficeSupported)
            {
                break
            }
        }

        $global:__OfficeChecked = $true
        if(!$OfficeInstalled)
        {
            return 2
        }
        elseif(!$OfficeSupported)
        {
            return 1
        }
        else
        {
            return 0
        }
#}
#    else
#    {
#        return 0
#    }
}

function Set-OfficeRegPath($Version,$OfficeRegPath)
{
    $global:__OfficeVersion = $Version
    $global:__OfficeRegPath = $OfficeRegPath
}

function Get-OfficeRegistrationKey()
{

    $OfficePathHash = @{"15.0"="*{*00-0000000FF1CE}";
	                    "16.0"="*{*00-0000000FF1CE}";
                        "14.0"="*{*00-0000000FF1CE}";
                         "12.0"="*{*000-0000000FF1CE}";
                         "11.0"="*{*-6000-11D3-8CFE-0150048383C9}"}
    $subKeys = Get-ChildItem -Path ($global:__OfficeRegPath + "\Registration")
    $OfficeRegistrationKey = $null
    foreach ($key in $subKeys)
    {
        if($key.Name -like $OfficePathHash.$global:__OfficeVersion)
        {
            $OfficeRegistrationKey = $key.PsChildName
            break
        }
    }
    return $OfficeRegistrationKey
}

function Get-OfficeProductID()
{PARAM ([string]$ProductNameWithEXE,[string]$OnlyProductName)

    Check-Office $ProductNameWithEXE $OnlyProductName | out-null
    $OfficeProductInfoPath = $global:__OfficeRegPath + "\Registration\" + $(Get-OfficeRegistrationKey)
    return (Get-item -path $OfficeProductInfoPath).getvalue("ProductID")
}

function CompareVersion($path,$SpecificVersion)
{
    $VersionInfo = [system.diagnostics.fileversioninfo]::getversioninfo((get-item "$path"))
    $Version = @($VersionInfo.productmajorpart,$VersionInfo.productminorpart,$VersionInfo.productbuildpart,$VersionInfo.productprivatepart)
    $result = $true

    foreach($i in 0..3)
    {
        if($Version[$i] -gt $SpecificVersion[$i])
        {
            $result = $true
            break
        }
        if($Version[$i] -lt $SpecificVersion[$i])
        {
            $result = $false
            break
        }
    }
    return $result
}

#Function to get the Product Name of Office 
# Ex: WordName; OutlookName; ExcelName; PowerPointName
# Eg: Get-OfficeProductName "WordName" "WinWord.exe" "Word"
function Get-OfficeProductName
{PARAM ( [string]$ProductName, [string]$ProductNameWithEXE,[string]$OnlyProductName) 
    Check-Office $ProductNameWithEXE $OnlyProductName | out-null
    $OfficeProductInfoPath = $global:__OfficeRegPath + "\Registration\" + $(Get-OfficeRegistrationKey)
    return (Get-item -path $OfficeProductInfoPath).getvalue($ProductName)
}

#Determine whether outlook is running or not
Function Is_outlookRunning
{
 $result = (Get-Process |?{$_.ProcessName  -ieq "Outlook"}) 
 if ($result) {
  return  $True
 }
 else{
  return  $False
 }
}

#Determine whether a process is running or not
Function IsRunning
{    param ($processName)

    $result = (Get-Process |?{$_.ProcessName  -ieq $processName }) 
	if ($result) 
	{
		return  $True
	}
	else
	{
		return  $False
	}
}


# Determine whether there is an open Office app
Function IsAnyOfficeAppRunning
{    param ($processNames)
   
   foreach ($processName in $processNames)
   {
    $result = (Get-Process |?{$_.ProcessName  -ieq $processName }) 
	if ($result) 
	{
		return  $True
	}
  }   
return  $False	
}

# Determine whether there is any open Office RT apps
Function IsAnyOfficeRTAppRunning
{    #param ($processNames)
   $processNames=@("Winword","excel", "oneNote", "powerpnt")
   foreach ($processName in $processNames)
   {
    $result = (Get-Process |?{$_.ProcessName  -ieq $processName }) 
	if ($result) 
	{
		return  $True
	}
  }   
return  $False	
}

#Stop outlook process if it's running
Function StopOutlook
{
	if (Is_outlookRunning)
		{
			Stop-Process -Name outlook -Force
			While (Is_outlookRunning) 
			{}
		}
}

#Stop process if it's running
Function StopProcess
{    
	param ($processName)

	if (IsRunning $processName)
		{
			Stop-Process -Name $processName -Force
			While (IsRunning $processName) 
			{}
		}
}

# to check whether TS is running or Detecting Additional Problems, only works if the pack is run as elevated
function isRunningDetectingAdditionalProblems($packName = "already.txt"){	
	# drop a file in the current folder so that it can be checked later
	"once" > ".\$packName"
	
	# pop-msg "stop here"
	$p1 = dir ($env:windir+"\Temp") |  Sort-Object LastWriteTime -descending # sorting the items of directory in descending order
	$p1 = $p1 |  where { $_.Mode -match "d" } # search for directories only, cause we aren’t interested in files

	# only works if the pack is run as elevated
	$dir1 = (($env:windir+"\Temp")+"\"+$p1[0].Name)
	$dir2 = (($env:windir+"\Temp")+"\"+$p1[1].Name)

	if( (test-path "$dir1\$packName") -and (test-path "$dir2\$packName") ){
		# check if 2 files are present or not, if the both files are present then its a load back
		return $true
	}	
	return $false	
}

function ispostbackOnWin8($packName){
	[string] $path1 = (Get-Location -PSProvider FileSystem).ProviderPath	
	[string] $path1 = $path1 + "\$packName"	
	if(test-path $path1){
		# del $path1 -force
		# the file is already there so this must be detecting additional problem		
		return $true
	}
	"once" > $path1 	
	return $false
}

# Function: To determine whether the specified version of office MSI is installed or not
# Argument:
#	1. $OffVer: Office version in the formats- Office15, Office14,Office12 and Office11
# Return:
#	$true if the specified version of office MSI is installed, false otherwise
# Usage example: isOfficeMSI_Installed "Office15"
function isOfficeMSI_Installed 
{
	Param ([string] $OffVer)
		
	$OfficeMSI_Paths = @("$env:CommonProgramFiles\microsoft Shared\$OffVer\mso.dll","${env:CommonProgramFiles(x86)}\microsoft Shared\$OffVer\mso.dll")
	foreach($OfficeMSI_Path in $OfficeMSI_Paths)
	{                
		if(Test-Path $OfficeMSI_Path)
		{
		 return $true
		}	
	}

	return $false
}

# Function: To determine whether office15/14 c2r is installed or not
# Argument:
#	1. $OffVer: Office version in the formats- Office15, Office14. No c2r suite before office 14 
# Return:
#	$true if office15/14 c2r is installed, false otherwise
# Usage example: isOfficeC2R_Installed "Office15"


function isOfficeC2R_Installed
{ 
	Param ([string] $OffVer)

	switch ($OffVer)
	{
	  "Office16" 
		{
			$Office16C2R_Paths = @("$env:ProgramFiles\Microsoft Office\root\Office16\c2r32.dll","$env:ProgramFiles\Microsoft Office\root\Office16\c2r64.dll")
			foreach($Office16C2R_Path in $Office16C2R_Paths)
			{                
				if(Test-Path $Office16C2R_Path)
				{
				 return $true
				}	
			}	
			return $false
		}
		"Office15" 
		{
			$Office15C2R_Paths = @("$env:ProgramFiles\Microsoft Office 15\root\Office15\c2r32.dll","$env:ProgramFiles\Microsoft Office 15\root\Office15\c2r64.dll")
			foreach($Office15C2R_Path in $Office15C2R_Paths)
			{                
				if(Test-Path $Office15C2R_Path)
				{
				 return $true
				}	
			}	
			return $false
		}
		"Office14" 
		{
			$Office14C2R_Paths = @("$env:ProgramFiles\Common Files\microsoft shared\Virtualization Handler\CVH.EXE","${env:CommonProgramFiles(x86)}\Common Files\microsoft shared\Virtualization Handler\CVH.EXE")
			foreach($Office14C2R_Path in $Office14C2R_Paths)
			{                
				if(Test-Path $Office14C2R_Path)
				{
				 return $true
				}	
			}	
			return $false
		}
	 }
}


# Function: To determine whether MS Office (c2r or msi, off15 or 14 or 12 or 11) is installed
# Argument:
#	1. No Arg
# Return:
#	$true if office is installed, false otherwise
# Usage example: isOfficeInstalled
Function isOfficeInstalled
{
 $O16 = ((isOfficeMSI_Installed "Office16") -OR (isOfficeC2R_Installed "Office16"))
 $O15 = ((isOfficeMSI_Installed "Office15") -OR (isOfficeC2R_Installed "Office15"))
 $O14 = ((isOfficeMSI_Installed "Office14") -OR (isOfficeC2R_Installed "Office14"))
 $O12 = (isOfficeMSI_Installed "Office12") 
 $O11 = (isOfficeMSI_Installed "Office11") 
 
 return (($O15) -or ($O14) -or ($O12 ) -or ($O11) -or ($O16) ) 		
}

# SIG # Begin signature block
# MIIayQYJKoZIhvcNAQcCoIIaujCCGrYCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU3H6Dt5/3rUUecQb/kWgmBpkv
# zlWgghWCMIIEwzCCA6ugAwIBAgITMwAAAG9lLVhtBxFGKAAAAAAAbzANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTUwMzIwMTczMjAy
# WhcNMTYwNjIwMTczMjAyWjCBszELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjENMAsGA1UECxMETU9QUjEnMCUGA1UECxMebkNpcGhlciBEU0UgRVNO
# OkMwRjQtMzA4Ni1ERUY4MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAz+ZtzcEqza6o
# XtiVTy0DQ0dzO7hC0tBXmt32UzZ31YhFJGrIq9Bm6YvFqg+e8oNGtirJ2DbG9KD/
# EW9m8F4UGbKxZ/jxXpSGqo4lr/g1E/2CL8c4XlPAdhzF03k7sGPrT5OaBfCiF3Hc
# xgyW0wAFLkxtWLN/tCwkcHuWaSxsingJbUmZjjo+ZpWPT394G2B7V8lR9EttUcM0
# t/g6CtYR38M6pR6gONzrrar4Q8SDmo2XNAM0BBrvrVQ2pNQaLP3DbvB45ynxuUTA
# cbQvxBCLDPc2Ynn9B1d96gV8TJ9OMD8nUDhmBrtdqD7FkNvfPHZWrZUgNFNy7WlZ
# bvBUH0DVOQIDAQABo4IBCTCCAQUwHQYDVR0OBBYEFPKmSSl4fFdwUmLP7ay3eyA0
# R9z9MB8GA1UdIwQYMBaAFCM0+NlSRnAK7UD7dvuzK7DDNbMPMFQGA1UdHwRNMEsw
# SaBHoEWGQ2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3Rz
# L01pY3Jvc29mdFRpbWVTdGFtcFBDQS5jcmwwWAYIKwYBBQUHAQEETDBKMEgGCCsG
# AQUFBzAChjxodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY3Jv
# c29mdFRpbWVTdGFtcFBDQS5jcnQwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZI
# hvcNAQEFBQADggEBAI2zTLbY7A2Hhhle5ADnl7jVz0wKPL33VdP08KCvVXKcI5e5
# girHFgrFJxNZ0NowK4hCulID5l7JJWgnJ41kp235t5pqqz6sQtAeJCbMVK/2kIFr
# Hq1Dnxt7EFdqMjYxokRoAZhaKxK0iTH2TAyuFTy3JCRdu/98U0yExA3NRnd+Kcqf
# skZigrQ0x/USaVytec0x7ulHjvj8U/PkApBRa876neOFv1mAWRDVZ6NMpvLkoLTY
# wTqhakimiM5w9qmc3vNTkz1wcQD/vut8/P8IYw9LUVmrFRmQdB7/u72qNZs9nvMQ
# FNV69h/W4nXzknQNrRbZEs+hm63SEuoAOyMVDM8wggTsMIID1KADAgECAhMzAAAB
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
# 9wFlb4kLfchpyOZu6qeXzjEp/w7FW1zYTRuh2Povnj8uVRZryROj/TGCBLEwggSt
# AgEBMIGQMHkxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
# VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xIzAh
# BgNVBAMTGk1pY3Jvc29mdCBDb2RlIFNpZ25pbmcgUENBAhMzAAABCix5rtd5e6as
# AAEAAAEKMAkGBSsOAwIaBQCggcowGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQw
# HAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFLOJ
# IxVosbOOJnujW//MdZfFoKs+MGoGCisGAQQBgjcCAQwxXDBaoECAPgBPAGYAZgBp
# AGMAZQBVAG4AaQBuAHMAdABhAGwAbABlAHIAXwBDAEwAXwBPAGYAZgBpAGMAZQAu
# AHAAcwAxoRaAFGh0dHA6Ly9taWNyb3NvZnQuY29tMA0GCSqGSIb3DQEBAQUABIIB
# AHQFZSF1D8s2/JSVd2iAxpknCnP22vuLviEirjYeQAevfMObjkDJMpZqaypw3+Lp
# mIU+NmlHZfvGZdsVMXTxTXDhE3gzajKxUIrO8zQFjsbyYWbgAJtni5dEPy9dR93v
# UQyqKFoU6UhTZ3Cas562lcgsZjkPRYLz/mjHdD5aJfEepWciV0aR3hwscq+e6V63
# ncv2Qp/npiaKJOOpj6IQA+Y0028nwUkSVLmhi7G8B8GwQJghc4M9oNGqOsUsbUpY
# oUpwRCkL1jSx25wAUkbGv/9tMKZvpcfaMSfggfVYb31NE/itknswm+/WUSi08rxi
# IBEqYe1+qTWCO9rAUpzXs1ehggIoMIICJAYJKoZIhvcNAQkGMYICFTCCAhECAQEw
# gY4wdzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcT
# B1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEhMB8GA1UE
# AxMYTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBAhMzAAAAb2UtWG0HEUYoAAAAAABv
# MAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3
# DQEJBTEPFw0xNTA3MjAyMTQ1MzBaMCMGCSqGSIb3DQEJBDEWBBTH1Rj4Pq6+M6c1
# qvdCIb2XcV3ACTANBgkqhkiG9w0BAQUFAASCAQAO9WTdIK0+u67ssPMUOWmHZuJr
# Br5f8Io6AB5UlbmmmP3eBkAatEvijbkTWhTXGEANbBMRyVy15kPS/B619mTZTXDx
# ZzcPs3oAF+AoTBNNvSkXFPqBomsC6TJpc9lMmQu5PLSvJQzGYrSOI/X46z8p08xU
# Xo/gbVnle959ggw86TnwkO00pJqgihHc1DwiS7hQXWTpbSwTRm+1fLBlDx8tVqVo
# Ps6rr9oLpI84qyf8lWyOr5SiFxbr1LF3e2iD33drg8beKecMTYXCY4RbKx/43NfE
# ygvc0/IZcqMWCd7FCh1OBKBeIuFjRAjkiQd6Wx4QkCmT9Z7WIrbHK+cIqUZ3
# SIG # End signature block
