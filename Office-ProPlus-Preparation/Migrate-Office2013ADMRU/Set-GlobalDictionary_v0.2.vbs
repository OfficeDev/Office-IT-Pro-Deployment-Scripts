' Script to set global dictionary

'The following functionality has been implemented:
'	Logging - Default folder for logs and backup files is %appdata%\Microsoft\OfficeSettingMigration\
'	Capturing debug information
'	Backup of Roaming Dictionary and Uproof\Custom Dictionary
'	Merging contents of Roaming Dictionary with Uproof\Custom Dictionary and ensuring that there are only unique entries
'	Updating registry to set Uproof\Custom disctionary as the default


'all variable names must be explicitly declared
Option Explicit

'~~~~~~ declare variables ~~~~~~
Dim WSHShell, logFileFolder, fso, backupFolderPath, logFileName, logFilePath, logFile
Dim logMsg, strComputer, objWMIService, osState, osVal
Dim roamingCustomDicPath, strOffice15RoamingFolder, objOffice15RoamingFolder, roamingCustomDicBackupPath
Dim strUproofFile, uproofCustomDicBackupPath
Dim oFile, strLine, dictUCustom, roamFile, fso2, oReg, strKeyPath, arrValueNames, arrValueTypes, ctr
Dim data, ret, err, strTime, dictRoamingCustom, dictElem, dictItemCount, oReg2

'~~~~~~ initialize global variables ~~~~~~

roamingCustomDicPath = Empty
Const HKEY_CURRENT_USER = &H80000001
dictItemCount = 0

'~~~~~~ initialize logging sub ~~~~~~
Set WSHShell = WScript.CreateObject("WScript.Shell")
logFileFolder = WSHShell.ExpandEnvironmentStrings("%APPDATA%")
logFileFolder = logFileFolder & "\Microsoft\OfficeSettingMigration\"

Set fso = CreateObject("Scripting.FileSystemObject")
if NOT (fso.FolderExists(logFileFolder)) then
	fso.CreateFolder(logFileFolder)
end if

backupFolderPath = logFileFolder

strTime = Now
strTime = Replace(strTime," ","_")
strTime = Replace(strTime,"/","_")
strTime = Replace(strTime,":","_")

logFileName = "Set-GlobalDictionary_" & strTime & ".log"
logFilePath = logFileFolder & logFileName

Set logFile = fso.CreateTextFile(logFilePath,True)
logFile.Close

'~~~~~~ log header ~~~~~~
logMsg = "++++Start Office Dictionary Swap Script"
Logger logFilePath, 1, logMsg, 1

logMsg = "Computername: " & WSHShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
Logger logFilePath, 1, logMsg, 1

logMsg = "UserName: " & WSHShell.ExpandEnvironmentStrings("%USERNAME%")
Logger logFilePath, 1, logMsg, 1

logMsg = "UserDomain: " & WSHShell.ExpandEnvironmentStrings("%USERDOMAIN%")
Logger logFilePath, 1, logMsg, 1

logMsg = "Processor Architecture: " & WSHShell.ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%")
Logger logFilePath, 1, logMsg, 1

strComputer = "."
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set osState = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem")

For Each osVal in osState
    logMsg =  "Boot Device: " & osVal.BootDevice
	Logger logFilePath, 1, logMsg, 1	
    logMsg =  "Build Number: " & osVal.BuildNumber
	Logger logFilePath, 1, logMsg, 1
    logMsg =  "Build Type: " & osVal.BuildType
	Logger logFilePath, 1, logMsg, 1
    logMsg =  "Caption: " & osVal.Caption
	Logger logFilePath, 1, logMsg, 1
    logMsg =  "Code Set: " & osVal.CodeSet
	Logger logFilePath, 1, logMsg, 1
    logMsg =  "Country Code: " & osVal.CountryCode
	Logger logFilePath, 1, logMsg, 1
    logMsg =  "Debug: " & osVal.Debug
	Logger logFilePath, 1, logMsg, 1
    logMsg =  "Encryption Level: " & osVal.EncryptionLevel
	Logger logFilePath, 1, logMsg, 1
    logMsg =  "Licensed Users: " & osVal.NumberOfLicensedUsers
	Logger logFilePath, 1, logMsg, 1
    logMsg =  "Organization: " & osVal.Organization
	Logger logFilePath, 1, logMsg, 1
    logMsg =  "OS Language: " & osVal.OSLanguage
	Logger logFilePath, 1, logMsg, 1
    logMsg =  "OS Product Suite: " & osVal.OSProductSuite
	Logger logFilePath, 1, logMsg, 1
    logMsg =  "OS Type: " & osVal.OSType
	Logger logFilePath, 1, logMsg, 1
    logMsg =  "Primary: " & osVal.Primary
	Logger logFilePath, 1, logMsg, 1
    logMsg =  "Registered User: " & osVal.RegisteredUser
	Logger logFilePath, 1, logMsg, 1
    logMsg =  "Serial Number: " & osVal.SerialNumber
	Logger logFilePath, 1, logMsg, 1
    logMsg =  "Version: " & osVal.Version
	Logger logFilePath, 1, logMsg, 1
Next

logMsg =  "++++Start Office Dictionary Swap Tasks"
Logger logFilePath, 1, logMsg, 1

'~~~~~~ get the path of all RoamingCustom.dic references in the "Proofing Tools\1.0\Custom Dictionaries" hive ~~~~~~
Set dictRoamingCustom = CreateObject("Scripting.Dictionary")
strComputer = "."
const REG_SZ = 1
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
Set oReg2=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "Software\Microsoft\Shared Tools\Proofing Tools\1.0\Custom Dictionaries"
oReg.EnumValues HKEY_CURRENT_USER, strKeyPath, arrValueNames, arrValueTypes

'WScript.Echo UBound(arrValueNames)
if Not IsNull(arrValueNames) then
	For ctr=0 To UBound(arrValueNames)		
		if arrValueTypes(ctr) = REG_SZ then	
			oReg2.GetStringValue HKEY_CURRENT_USER, strKeyPath, arrValueNames(ctr), data
			'WScript.Echo "Value Name: " & arrValueNames(ctr) & " Value Type: " & arrValueTypes(ctr) & " Data: " & data
			if InStr(1,data,"RoamingCustom.dic",vbTextCompare)>0 then
				'WScript.Echo "Value Name: " & arrValueNames(ctr) & " Value Type: " & arrValueTypes(ctr) & " Data: " & data
				if NOT dictRoamingCustom.Exists(data) then
					if fso.FileExists(data) then
						dictRoamingCustom.Add data, data
						logMsg =  "RoamingCustom.dic Found at: " & data
						Logger logFilePath, 1, logMsg, 1
						dictItemCount = dictItemCount + 1
					else
						logMsg =  "RoamingCustom.dic entry in registry NOT Found at: " & data
						Logger logFilePath, 1, logMsg, 2
					end if
				end if
			end if
		end if
	Next
end if


'~~~~~~ backup RoamingCustom.dic ~~~~~~

if dictItemCount < 1 then
	logMsg =  "RoamingCustom.dic file not found!"
	Logger logFilePath, 1, logMsg, 2
else
	logMsg =  "Count of RoamingCustom.dic files found: " & dictItemCount
	Logger logFilePath, 1, logMsg, 1
	logMsg =  "++Initiating backup of RoamingCustom.dic"
	Logger logFilePath, 1, logMsg, 1
end if

ctr = 1
For Each dictElem in dictRoamingCustom
	roamingCustomDicPath = dictElem
	logMsg =  "RoamingCustom.dic path: " & roamingCustomDicPath
	Logger logFilePath, 1, logMsg, 1

	roamingCustomDicBackupPath = backupFolderPath & "Bk_" & ctr & "__" & strTime & "_" & fso.GetFileName(roamingCustomDicPath)
	ctr = ctr + 1
	logMsg =  "RoamingCustom.dic backup path: " & roamingCustomDicBackupPath
	Logger logFilePath, 1, logMsg, 1

	fso.CopyFile roamingCustomDicPath, roamingCustomDicBackupPath

	if fso.FileExists(roamingCustomDicBackupPath) then
		logMsg =  "RoamingCustom.dic successfully backed up to : " & roamingCustomDicBackupPath
		Logger logFilePath, 1, logMsg, 1
	else
		logMsg =  "Backup of RoamingCustom.dic failed. Terminating Script"
		Logger logFilePath, 1, logMsg, 3
		Wscript.Quit
	end if
Next

'~~~~~~ backup Uproof\Custom.dic ~~~~~~
strUproofFile = WSHShell.ExpandEnvironmentStrings("%APPDATA%") & "\Microsoft\UProof\CUSTOM.DIC"

if fso.FileExists(strUproofFile) then
	logMsg =  "Uproof\Custom.dic path: " & strUproofFile
	Logger logFilePath, 1, logMsg, 1
	
	uproofCustomDicBackupPath = backupFolderPath & "Bk_" & strTime & "_Uproof_CUSTOM.DIC"
	logMsg =  "Uproof\Custom.dic backup path: " & uproofCustomDicBackupPath
	Logger logFilePath, 1, logMsg, 1	
	
	fso.CopyFile strUproofFile, uproofCustomDicBackupPath
	
	if fso.FileExists(uproofCustomDicBackupPath) then
		logMsg =  "Uproof\Custom.dic successfully backed up to : " & uproofCustomDicBackupPath
		Logger logFilePath, 1, logMsg, 1
	else
		logMsg =  "Backup of Uproof\Custom.dic failed"
		Logger logFilePath, 1, logMsg, 3
	end if
else
	logMsg =  "Uproof\Custom.dic file not found. Terminating Script"
	Logger logFilePath, 1, logMsg, 3
	Wscript.Quit
end if

'~~~~~~ Merge the contents of RoamingCustom.dic with Uproof\Custom.dic ~~~~~~
logMsg =  "++Start merge of RoamingCustom.dic with Uproof\Custom.dic"
Logger logFilePath, 1, logMsg, 1

Const ForReading = 1
Const ForAppending = 8
Const AsUnicode = -1
Set dictUCustom = CreateObject("Scripting.Dictionary")
Set oFile = fso.OpenTextFile(uproofCustomDicBackupPath, ForReading, false, AsUnicode)

Do Until oFile.AtEndOfStream
	strLine = oFile.ReadLine
	if NOT dictUCustom.Exists(strLine) then
		dictUCustom.Add strLine, strLine
		'Wscript.Echo strLine
	end if
Loop
oFile.Close
Set oFile = Nothing 

Set fso2 = CreateObject("Scripting.FileSystemObject")
Set oFile = fso2.OpenTextFile(strUproofFile, ForAppending, true, AsUnicode)

For Each dictElem in dictRoamingCustom
	roamingCustomDicPath = dictElem
	Set roamFile = fso.OpenTextFile(roamingCustomDicPath, ForReading, false, AsUnicode)
	Do Until roamFile.AtEndOfStream
		strLine = roamFile.ReadLine
		if NOT dictUCustom.Exists(strLine) then
			oFile.WriteLine strLine
		end if
	Loop
	logMsg =  "Contents of [" &  roamingCustomDicPath & "] merged with Uproof\Custom.dic"
	Logger logFilePath, 1, logMsg, 1
	roamFile.Close
	Set roamFile = Nothing
Next

oFile.Close
Set oFile = Nothing



'~~~~~~ Modify registry to replace entry of RoamingCustom.dic with CUSTOM.DIC ~~~~~~
'strComputer = "."
'Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
'strKeyPath = "Software\Microsoft\Shared Tools\Proofing Tools\1.0\Custom Dictionaries"

logMsg =  "++Modify registry to replace entry/entries of RoamingCustom.dic with CUSTOM.DIC"
Logger logFilePath, 1, logMsg, 1

oReg.EnumValues HKEY_CURRENT_USER, strKeyPath, arrValueNames, arrValueTypes
if Not IsNull(arrValueNames) then
	For ctr=0 To UBound(arrValueNames)		
		if arrValueTypes(ctr) = REG_SZ then	
			oReg.GetStringValue HKEY_CURRENT_USER, strKeyPath, arrValueNames(ctr), data
			'WScript.Echo "Value Name: " & arrValueNames(ctr) & " Value Type: " & arrValueTypes(ctr) & " Data: " & data
			if InStr(1,data,"RoamingCustom.dic",vbTextCompare)>0 then
				'WScript.Echo "Value Name: " & arrValueNames(ctr) & " Value Type: " & arrValueTypes(ctr) & " Data: " & data
				ret = oReg.SetStringValue(HKEY_CURRENT_USER, strKeyPath, arrValueNames(ctr), "CUSTOM.DIC")
				    if ret = 0 then 
						logMsg =  "Registry Updated."
						Logger logFilePath, 1, logMsg, 1
						logMsg =  "ValueName: Software\Microsoft\Shared Tools\Proofing Tools\1.0\Custom Dictionaries\" & arrValueNames(ctr)
						Logger logFilePath, 1, logMsg, 1
						logMsg =  "Old Value: " & data & " Updated Value: " & "CUSTOM.DIC"
						Logger logFilePath, 1, logMsg, 1						 
					Else
						logMsg =  "Registry Failed to Update."
						Logger logFilePath, 1, logMsg, 3
						logMsg =  "ValueName: Software\Microsoft\Shared Tools\Proofing Tools\1.0\Custom Dictionaries\" & arrValueNames(ctr)
						Logger logFilePath, 1, logMsg, 1
						logMsg =  "Terminating Script."
						Logger logFilePath, 1, logMsg, 1
						Wscript.Quit
					End If
			end if
		end if
	Next
end if
	
logMsg =  "Script Completed!!!"
Logger logFilePath, 1, logMsg, 1
	

'~~~~~~ Logging Subroutine ~~~~~~
'inputs: log file name : string
'thread id : int
'message : string
'severity : 1=info, 2=warning, 3=error, 4=verbose
Sub Logger(strRemLogFile, nPhase, Msg, Svrty) 
		 
	Dim objLogFile, arrSvrty, SvrtyCheck, i 
			 
	Const ForAppending = 8 
			 
	arrSvrty = Array(1,2,3) 
	SvrtyCheck = "" 
	For i = 0 To UBound(arrSvrty) 
		If CInt(arrSvrty(i)) = CInt(Svrty) Then 
			SvrtyCheck = "OK" 
			Exit For 
		End If 
	Next 
 
	If SvrtyCheck <> "OK" Then 
		'MsgBox "Unrecognised Severity in Logger Sub" & Chr(13) & Chr(13), 16, WScript.ScriptName 
	End If 
			 

	Set objLogFile = fso.OpenTextFile(strRemLogFile, ForAppending, True) 
	
	
	'MsgBox "<![LOG[" & Msg & "]LOG]!>" & "<time=" & Chr(34) & Time & ".000-00" & Chr(34) & " date=" & Chr(34) _ 
'	        & Month(Now) & "-" & Day(Now) & "-" & Year(Now) & Chr(34) & " component=" & Chr(34) & WScript.ScriptName _ 
'	        & Chr(34) & " context=" & Chr(34) & Chr(34) & " type=" & Chr(34) & Svrty & Chr(34) & " thread=" _ 
'	        & Chr(34) & nPhase & Chr(34) & " file=" & Chr(34) & WScript.ScriptFullName & Chr(34) & ">"
		
		objLogFile.WriteLine "<![LOG[" & Msg & "]LOG]!>" & "<time=" & Chr(34) & FormatDateTime (Now, 4) & ":" & Left(Right(FormatDateTime (Now),5),2) &".000-00" & Chr(34) & " date=" & Chr(34) _ 
		& Month(Now) & "-" & Day(Now) & "-" & Year(Now) & Chr(34) & " component=" & Chr(34) & WScript.ScriptName _ 
		& Chr(34) & " context=" & Chr(34) & Chr(34) & " type=" & Chr(34) & Svrty & Chr(34) & " thread=" _ 
		& Chr(34) & nPhase & Chr(34) & " file=" & Chr(34) & WScript.ScriptFullName & Chr(34) & ">"  
		objLogFile.Close 
  
	Set objLogFile = Nothing 
	   
End Sub

