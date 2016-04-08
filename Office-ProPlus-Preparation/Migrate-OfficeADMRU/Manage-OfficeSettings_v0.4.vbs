'Script to manage Office 2013 settings (Custom Dictionary and MRU) across Identity Federation

'The following functionality has been implemented:
'Logging - Default folder for logs and backup files is %appdata%\Microsoft\OfficeSettingMigration\
'Log debug header - key environment variables, user settings and computer settings
'Log current SignInOptions
'Backup RoamingCustom.dic
'Get initial set of identities present on the computer
'Backup MRUs for each Office app and merge the result to create a single reg file (AD_MRURegBackup.reg)
'Backup roaming settings and customizations (RoamingADSettingsBackup.reg)
'Log the list of Office Apps that have MRU
'Launch Apps and wait for AAD Auth (as signaled by presence of OrgId)
'Word is launched first.
'Excel, PowerPoint, Access and Publisher are launched if MRU exists for these apps
'A reg file (OrgId_MRURegBackup.reg) is created with MRUs and OrgId settings. This file is merged with the registry to restore MRUs

'all variable names must be explicitly declared
Option Explicit


'On Error Resume Next

'~~~~~~ declare variables ~~~~~~
Dim SignInOptionsLoc, SignInOptionsValue
Dim WSHShell, oReg, strComputer, strKeyPath, strValueName, strStringValues
Dim logFilePath, logMsg, logFile, logFileFolder, fso, logFileName
Dim objWMIService, osState, osVal
Dim strOffice15RoamingFolder, objOffice15RoamingFolder, roamingCustomDicPath, roamingCustomDicBackupPath
Dim backupFolderPath, OsType
Dim strIdentitiesPath, regIdentity, dictIdentity, hiveCounter, dictItems, dictIter
Dim identity_AD, identity_OrgId, identity_AD_Count, identity_OrgId_Count
Dim strMRURoot, regMRU, dictBackupMRU
Dim strADKey, strOrgIdKey, tmpArr, ctr, tmpStr
Dim strRoamingRoot, regRoaming, dictBackupRoaming
Dim objShell, strCommand, dictApps, strApp, dictAppsExePath, officeInstallRoot, objOfficeInstallRoot, dictElem
Dim process, aadAuth, regMRUOrgID, arrSubKeys, subkey
Dim sourceFile, strLine, fso2, targetFile, strReplaceLine

'~~~~~~ initialize global variables ~~~~~~

roamingCustomDicPath = Empty
identity_AD = 0
identity_OrgId = 0
aadAuth = 0

'~~~~~~ pre execution validation ~~~~~~


'~~~~~~ initialize logging sub ~~~~~~
Set WSHShell = WScript.CreateObject("WScript.Shell")
logFileFolder = WSHShell.ExpandEnvironmentStrings("%APPDATA%")
logFileFolder = logFileFolder & "\Microsoft\OfficeSettingMigration\"

Set fso = CreateObject("Scripting.FileSystemObject")
if NOT (fso.FolderExists(logFileFolder)) then
	fso.CreateFolder(logFileFolder)
end if

backupFolderPath = logFileFolder

logFileName = "Manage-OfficeSettings_" & Now & ".log"
logFileName = Replace(logFileName," ","_")
logFileName = Replace(logFileName,"/","_")
logFileName = Replace(logFileName,":","_")

logFilePath = logFileFolder & logFileName

Set logFile = fso.CreateTextFile(logFilePath,True)
logFile.Close

'~~~~~~ log header ~~~~~~
logMsg = "++Start Office Settings Migration Script"
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

logMsg =  "++Start Office Settings Backup Tasks"
Logger logFilePath, 1, logMsg, 1

'~~~~~~ get current SignInOptions ~~~~~~
Set WSHShell = WScript.CreateObject("WScript.Shell")
SignInOptionsLoc = "HKCU\Software\Microsoft\Office\15.0\Common\SignIn\SignInOptions"
SignInOptionsValue = WSHShell.RegRead(SignInOptionsLoc)

logMsg =  "Initial SignInOptionsValue: " & SignInOptionsValue
Logger logFilePath, 1, logMsg, 1

'~~~~~~ backup RoamingCustom.dic ~~~~~~
Set fso = CreateObject("Scripting.FileSystemObject")
strOffice15RoamingFolder = WSHShell.ExpandEnvironmentStrings("%APPDATA%") & "\Microsoft\Office\15.0\"
Set objOffice15RoamingFolder = fso.GetFolder(strOffice15RoamingFolder)
Call ShowSubfolders (objOffice15RoamingFolder)

if IsEmpty(roamingCustomDicPath) then
	logMsg =  "RoamingCustom.dic file not found!"
	Logger logFilePath, 1, logMsg, 3
else
	logMsg =  "RoamingCustom.dic path: " & roamingCustomDicPath
	Logger logFilePath, 1, logMsg, 1

	roamingCustomDicBackupPath = backupFolderPath & fso.GetFileName(roamingCustomDicPath)
	logMsg =  "RoamingCustom.dic backup path: " & roamingCustomDicBackupPath
	Logger logFilePath, 1, logMsg, 1

	fso.CopyFile roamingCustomDicPath, roamingCustomDicBackupPath

	if fso.FileExists(roamingCustomDicBackupPath) then
		logMsg =  "RoamingCustom.dic successfully backed up to : " & roamingCustomDicBackupPath
		Logger logFilePath, 1, logMsg, 1
	else
		logMsg =  "Backup of RoamingCustom.dic failed"
		Logger logFilePath, 1, logMsg, 3
	end if
end if

'~~~~~~ Check identity ~~~~~~
'Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_CURRENT_USER = &H80000001
strIdentitiesPath = "SOFTWARE\Microsoft\Office\15.0\Common\Identity\Identities"
Set regIdentity = GetObject("winmgmts://./root/default:StdRegProv")
Set dictIdentity = CreateObject("Scripting.Dictionary")
hiveCounter = 1
EnumerateKeys HKEY_CURRENT_USER, strIdentitiesPath

identity_AD_Count = 0
identity_OrgId_Count = 0
dictItems = dictIdentity.Items
For dictIter = 0 To UBound(dictItems)
	'Wscript.Echo dictItems(dictIter)
	if InStr(1,dictItems(dictIter),"_AD",vbTextCompare)>0 then
		identity_AD = 1
		identity_AD_Count = identity_AD_Count + 1
		logMsg =  "_AD identity detected: " & dictItems(dictIter)
		Logger logFilePath, 1, logMsg, 1
	end if
	if InStr(1,dictItems(dictIter),"OrgId_",vbTextCompare)>0 then
		identity_OrgId = 1
		identity_OrgId_Count = identity_OrgId_Count + 1
		logMsg =  "OrgId_ identity detected: " & dictItems(dictIter)
		Logger logFilePath, 1, logMsg, 1
	end if
Next


'~~~~~~ backup MRU ~~~~~~
' backup: HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\15.0\**App**\User MRU\AD_**\
strMRURoot = "SOFTWARE\Microsoft\Office\15.0"
Set regMRU = GetObject("winmgmts://./root/default:StdRegProv")
Set dictBackupMRU = CreateObject("Scripting.Dictionary")
hiveCounter = 1
EnumerateMRU HKEY_CURRENT_USER, strMRURoot

dictItems = dictBackupMRU.Items
For dictIter = 0 To UBound(dictItems)
	'Wscript.Echo dictItems(dictIter)
	tmpArr = Split(dictItems(dictIter), "\", -1, vbTextCompare)
	For Each tmpStr in tmpArr
		'Wscript.Echo tmpStr
		if	InStr(1, tmpStr, "AD_", vbTextCompare)>0 then
			strADKey = tmpStr
			logMsg =  "AD_ Key obtained: " & strADKey
			Logger logFilePath, 1, logMsg, 1
			Exit For
		end if
	Next
	if NOT IsEmpty(strADKey) then
		Exit For
	end if
Next

BackupMRU backupFolderPath & "AD_MRURegBackup.reg"

if fso.FileExists(backupFolderPath & "AD_MRURegBackup.reg") = True then
	logMsg =  "MRU Registry Successfully Backed up to: " & backupFolderPath & "AD_MRURegBackup.reg"
	Logger logFilePath, 1, logMsg, 1
else
	logMsg =  "MRU Registry Failed to backed up"
	Logger logFilePath, 1, logMsg, 3
end if

'~~~~~~ backup roaming settings and customizations ~~~~~~
'backup: HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\15.0\Common\Roaming\Identities\**_AD
strRoamingRoot = "SOFTWARE\Microsoft\Office\15.0\Common\Roaming\Identities"
Set regRoaming = GetObject("winmgmts://./root/default:StdRegProv")
Set dictBackupRoaming = CreateObject("Scripting.Dictionary")
hiveCounter = 1
EnumerateRoaming HKEY_CURRENT_USER, strRoamingRoot

Set objShell = CreateObject("WScript.Shell")
dictItems = dictBackupRoaming.Items
For dictIter = 0 To UBound(dictItems)
	'Wscript.Echo dictItems(dictIter)
	logMsg =  "Roaming Identity found at: " & dictItems(dictIter)
	Logger logFilePath, 1, logMsg, 1
	strCommand = "REG EXPORT """ & dictItems(dictIter) & """ """ & backupFolderPath & "RoamingADSettingsBackup.reg"" /y"
	'Wscript.Echo strCommand
	objShell.Run strCommand, 0, True
	if fso.FileExists(backupFolderPath & "RoamingADSettingsBackup.reg") = True then
		logMsg =  "Roaming Identity Settings backed up at: " & backupFolderPath & "RoamingADSettingsBackup.reg"
		Logger logFilePath, 1, logMsg, 1
	else
		logMsg =  "Roaming Identity Settings Failed to back up"
		Logger logFilePath, 1, logMsg, 3
	end if
Next

'~~~~~~ get list of Office apps with MRU. Only these apps need to be launched ~~~~~~
Set dictApps = CreateObject("Scripting.Dictionary")
dictItems = dictBackupMRU.Items
For dictIter = 0 To UBound(dictItems)
	strApp = dictItems(dictIter)
	'Wscript.Echo strApp
	strApp = Replace(strApp, "HKCU\SOFTWARE\Microsoft\Office\15.0\", "", 1, -1, vbTextCompare)
	strApp = Left(strApp, InStr(1, strApp, "\", vbTextCompare)-1 )
	'Wscript.Echo strApp
	if NOT dictApps.Exists(strApp) then
		dictApps.Add strApp, strApp
	end if
Next

logMsg =  "The following " & UBound(dictApps.Items)+1 & " app(s) have MRUs. Only these app(s) need to be launched for AAD Migration" 
Logger logFilePath, 1, logMsg, 1
dictItems = dictApps.Items
For dictIter = 0 To UBound(dictItems)
	logMsg =  dictIter+1 & ") App: " & dictItems(dictIter)
	Logger logFilePath, 1, logMsg, 1
Next

'~~~~~~ get path to app executables ~~~~~~
Set dictAppsExePath = CreateObject("Scripting.Dictionary")
'officeInstallRoot = WSHShell.ExpandEnvironmentStrings("%ProgramFiles%") & "\Microsoft Office\"

'OsType = WSHShell.RegRead("HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\PROCESSOR_ARCHITECTURE")    
'If OsType = "x86" then  
'		officeInstallRoot = WSHShell.ExpandEnvironmentStrings("%PROGRAMFILES%")  
'elseif OsType = "AMD64" then  
'		officeInstallRoot = WSHShell.ExpandEnvironmentStrings("%PROGRAMFILES(x86)%")  
'end if 

officeInstallRoot = WSHShell.ExpandEnvironmentStrings("%SystemDrive%") & "\"
'Wscript.Echo officeInstallRoot
Set objOfficeInstallRoot =  fso.GetFolder(officeInstallRoot)
FindFileEXEPath objOfficeInstallRoot, "winword.exe"
FindFileEXEPath objOfficeInstallRoot, "excel.exe"
FindFileEXEPath objOfficeInstallRoot, "powerpnt.exe"
FindFileEXEPath objOfficeInstallRoot, "msaccess.exe"
FindFileEXEPath objOfficeInstallRoot, "mspub.exe"

logMsg =  "App Path:"
Logger logFilePath, 1, logMsg, 1
For Each dictElem in dictAppsExePath
	'Wscript.Echo dictElem & " : " & dictAppsExePath(dictElem)
	logMsg =  dictElem & " : " & dictAppsExePath(dictElem)
	Logger logFilePath, 1, logMsg, 1
Next

'~~~~~~ set SignInOptions = 0 ~~~~~~
'Const HKEY_CURRENT_USER = &H80000001

strComputer = "."
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "Software\Microsoft\Office\15.0\Common\SignIn"
strValueName = "SignInOptions"
strStringValues = "0"

oReg.SetDWORDValue HKEY_CURRENT_USER,strKeyPath,strValueName,strStringValues
SignInOptionsValue = WSHShell.RegRead(SignInOptionsLoc)

'Msgbox SignInOptionsValue
logMsg =  "Modified SignInOptionsValue: " & SignInOptionsValue
Logger logFilePath, 1, logMsg, 1


'~~~~~~ launch app and wait for AAD Auth ~~~~~~
logMsg =  "++ Initiating App Launch"
Logger logFilePath, 1, logMsg, 1

Set process = Nothing
if dictAppsExePath.Exists("winword.exe")=True then 'dictApps.Exists("Word")=True
	strCommand = """" & dictAppsExePath("winword.exe") & """"
	logMsg =  "Launching Word: " & strCommand
	Logger logFilePath, 1, logMsg, 1	
	'Wscript.Echo strCommand
	Set process = objShell.Exec(strCommand)
	'Wscript.Echo process.ProcessID
	logMsg =  "Word Running. Process ID: " & process.ProcessID
	Logger logFilePath, 1, logMsg, 1
end if

strIdentitiesPath = "SOFTWARE\Microsoft\Office\15.0\Common\Identity\Identities"
Set regIdentity = GetObject("winmgmts://./root/default:StdRegProv")

For ctr = 0 to 5
	logMsg =  "Waiting for AAD Auth (" & ctr*10 & " seconds)"
	Logger logFilePath, 1, logMsg, 1
	WScript.Sleep 10000
	
	' check identity
	Set dictIdentity = CreateObject("Scripting.Dictionary")
	hiveCounter = 1
	EnumerateKeys HKEY_CURRENT_USER, strIdentitiesPath

	identity_AD_Count = 0
	identity_OrgId_Count = 0
	dictItems = dictIdentity.Items
	For dictIter = 0 To UBound(dictItems)
		'Wscript.Echo dictItems(dictIter)
		if InStr(1,dictItems(dictIter),"_AD",vbTextCompare)>0 then
			identity_AD = 1
			identity_AD_Count = identity_AD_Count + 1
			logMsg =  "Checking Identity: _AD identity detected: " & dictItems(dictIter)
			Logger logFilePath, 1, logMsg, 1
		end if
		if InStr(1,dictItems(dictIter),"_OrgId",vbTextCompare)>0 then
			identity_OrgId = 1
			identity_OrgId_Count = identity_OrgId_Count + 1
			logMsg =  "Checking Identity: _OrgId identity detected: " & dictItems(dictIter)
			Logger logFilePath, 1, logMsg, 1
		end if
	Next
	
	if identity_OrgId_Count = 1 then
		aadAuth = 1
		Exit For
	end if	
Next

'close Word
if NOT process is Nothing then
	strCommand = "cmd.exe /c taskkill /pid " & process.ProcessID & " /f"
	'Wscript.Echo strCommand
	logMsg =  "Terminating Word: " & strCommand
	Logger logFilePath, 1, logMsg, 1
	Set process = objShell.Exec(strCommand)
end if

if identity_OrgId_Count = 0 then
	logMsg =  "AAD Auth Failed. Terminating further script steps"
	Logger logFilePath, 1, logMsg, 3
	Wscript.Quit
else
	'launch remaining apps
	Set process = Nothing
	if dictApps.Exists("Excel")=True And dictAppsExePath.Exists("excel.exe")=True then
		strCommand = """" & dictAppsExePath("excel.exe") & """"
		logMsg =  "Launching Excel: " & strCommand
		Logger logFilePath, 1, logMsg, 1			
		Set process = objShell.Exec(strCommand)
		logMsg =  "Excel Running. Process ID: " & process.ProcessID
		Logger logFilePath, 1, logMsg, 1
		WScript.Sleep 10000

		if NOT process is Nothing then
			strCommand = "cmd.exe /c taskkill /pid " & process.ProcessID & " /f"
			'Wscript.Echo strCommand
			logMsg =  "Terminating Excel: " & strCommand
			Logger logFilePath, 1, logMsg, 1
			Set process = objShell.Exec(strCommand)
		end if
	end if
	
	Set process = Nothing
	if dictApps.Exists("PowerPoint")=True And dictAppsExePath.Exists("powerpnt.exe")=True then
		strCommand = """" & dictAppsExePath("powerpnt.exe") & """"
		logMsg =  "Launching PowerPoint: " & strCommand
		Logger logFilePath, 1, logMsg, 1			
		Set process = objShell.Exec(strCommand)
		logMsg =  "PowerPoint Running. Process ID: " & process.ProcessID
		Logger logFilePath, 1, logMsg, 1
		WScript.Sleep 10000

		if NOT process is Nothing then
			strCommand = "cmd.exe /c taskkill /pid " & process.ProcessID & " /f"
			'Wscript.Echo strCommand
			logMsg =  "Terminating PowerPoint: " & strCommand
			Logger logFilePath, 1, logMsg, 1
			Set process = objShell.Exec(strCommand)
		end if
	end if
	
	Set process = Nothing
	if dictApps.Exists("Access")=True And dictAppsExePath.Exists("msaccess.exe")=True then
		strCommand = """" & dictAppsExePath("msaccess.exe") & """"
		logMsg =  "Launching Access: " & strCommand
		Logger logFilePath, 1, logMsg, 1			
		Set process = objShell.Exec(strCommand)
		logMsg =  "Access Running. Process ID: " & process.ProcessID
		Logger logFilePath, 1, logMsg, 1
		WScript.Sleep 10000

		if NOT process is Nothing then
			strCommand = "cmd.exe /c taskkill /pid " & process.ProcessID & " /f"
			'Wscript.Echo strCommand
			logMsg =  "Terminating Access: " & strCommand
			Logger logFilePath, 1, logMsg, 1
			Set process = objShell.Exec(strCommand)
		end if
	end if
	
	Set process = Nothing
	if dictApps.Exists("Publisher")=True And dictAppsExePath.Exists("mspub.exe")=True then
		strCommand = """" & dictAppsExePath("mspub.exe") & """"
		logMsg =  "Launching Publisher: " & strCommand
		Logger logFilePath, 1, logMsg, 1			
		Set process = objShell.Exec(strCommand)
		logMsg =  "Publisher Running. Process ID: " & process.ProcessID
		Logger logFilePath, 1, logMsg, 1
		WScript.Sleep 10000

		if NOT process is Nothing then
			strCommand = "cmd.exe /c taskkill /pid " & process.ProcessID & " /f"
			'Wscript.Echo strCommand
			logMsg =  "Terminating Publisher: " & strCommand
			Logger logFilePath, 1, logMsg, 1
			Set process = objShell.Exec(strCommand)
		end if
	end if
end if

'~~~~~~ restore MRU settings ~~~~~~
' get OrgId_ string
strMRURoot = "Software\Microsoft\Office\15.0\Word\User MRU"
Set regMRUOrgID = GetObject("winmgmts://./root/default:StdRegProv")

regMRUOrgID.EnumKey HKEY_CURRENT_USER, strMRURoot, arrSubKeys
If Not IsNull(arrSubKeys) Then
	For Each subkey In arrSubKeys
		'Wscript.Echo subkey
		if InStr(1,subkey,"OrgId_",vbTextCompare)>0 then
			strOrgIdKey = subkey			
			logMsg =  "OrgId_ Key obtained: " & strOrgIdKey
			Logger logFilePath, 1, logMsg, 1
			Exit For
		end if
	Next
End If

Const ForReading = 1
Const ForAppending = 8
Const AsUnicode = -1
Const ForWriting = 2

Set sourceFile = fso.OpenTextFile(backupFolderPath & "AD_MRURegBackup.reg", ForReading, false, AsUnicode)
Set fso2 = CreateObject("Scripting.FileSystemObject")
Set targetFile = fso2.OpenTextFile(backupFolderPath & "OrgId_MRURegBackup.reg", ForWriting, true, AsUnicode)

Do Until sourceFile.AtEndOfStream
	strLine = sourceFile.ReadLine
	if InStr(1,strLine,strADKey,vbTextCompare)>0 then
		strReplaceLine = Replace(strLine, strADKey, strOrgIdKey, 1, -1, vbTextCompare)
	else
		strReplaceLine = strLine
	end if
	targetFile.WriteLine strReplaceLine
Loop
sourceFile.Close
Set sourceFile = Nothing 
targetFile.Close
Set targetFile = Nothing 

logMsg =  "MRU File created with OrgID: " & backupFolderPath & "OrgId_MRURegBackup.reg"
Logger logFilePath, 1, logMsg, 1

strCommand = "REG IMPORT """ & backupFolderPath & "OrgId_MRURegBackup.reg"""
objShell.Run strCommand, 0, True

logMsg =  "MRU File with OrgID merged with registry" & backupFolderPath & "OrgId_MRURegBackup.reg"
Logger logFilePath, 1, logMsg, 1

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
			 
'   On Error Resume Next 

	Set objLogFile = FSO.OpenTextFile(strRemLogFile, ForAppending, True) 
	
	
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

'~~~~~~ get location of RoamingCustom.dic ~~~~~~
Sub ShowSubFolders(fFolder)
    Dim objFolder, colFiles, fso, objFile, Subfolder
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set objFolder = fso.GetFolder(fFolder.Path)
    Set colFiles = objFolder.Files
    For Each objFile in colFiles
        If UCase(fso.GetExtensionName(objFile.name)) = "DIC" Then
            'Wscript.Echo objFile.Name
			'uses global variable
			roamingCustomDicPath = objFile.Path
        End If
    Next

    For Each Subfolder in fFolder.SubFolders
        ShowSubFolders Subfolder
    Next
End Sub

'~~~~~~ recurse through the registry keys and store values in a dictionary ~~~~~~
Sub EnumerateKeys(hive, key)
	Dim arrSubKeys, subkey
	'WScript.Echo "EnumerateKeys: " & hiveCounter & "->" & key
	dictIdentity.Add hiveCounter, key
	hiveCounter = hiveCounter + 1
	regIdentity.EnumKey hive, key, arrSubKeys
	If Not IsNull(arrSubKeys) Then
		For Each subkey In arrSubKeys
			EnumerateKeys hive, key & "\" & subkey
		Next
	End If
End Sub

'~~~~~~ recurse through the registry keys and save MRU keys in a dictionary ~~~~~~
Sub EnumerateMRU(hive, key)
	Dim arrSubKeys, subkey
	'WScript.Echo key
	'hiveCounter = hiveCounter + 1
	
	if InStr(1,key,"User MRU",vbTextCompare)>0 And InStr(1,key,"AD_",vbTextCompare)>0 then
		'WScript.Echo key
		dictBackupMRU.Add hiveCounter, "HKCU\"&key
		hiveCounter = hiveCounter + 1
		Exit Sub
	end if
'	if InStr(1,key,"User MRU",vbTextCompare)>1 And InStr(1,key,"Place MRU",vbTextCompare)>1 And InStr(1,key,"AD_",vbTextCompare)>1 then
		'WScript.Echo key
'		dictBackupMRU.Add hiveCounter, key
'		hiveCounter = hiveCounter + 1
'		Exit Sub
'	end if
	
	regMRU.EnumKey hive, key, arrSubKeys
	If Not IsNull(arrSubKeys) Then
		For Each subkey In arrSubKeys
			EnumerateMRU hive, key & "\" & subkey
		Next
	End If
End Sub

'~~~~~~ backup MRU into a single reg file ~~~~~~
Sub BackupMRU(bkFilePath)
	Dim objBkShell, objBkFSO, objBkRegFile, strRegPath, objInputFile, strCommand
	Set objBkShell = CreateObject("WScript.Shell")
	Set objBkFSO = CreateObject("Scripting.FileSystemObject")
	Const intForReading = 1
	Const intUnicode = -1

	'Wscript.Echo bkFilePath
	Set objBkRegFile = objBkFSO.CreateTextFile(bkFilePath, True, True)
	objBkRegFile.WriteLine "Windows Registry Editor Version 5.00"

	dictItems = dictBackupMRU.Items
	For dictIter = 0 To UBound(dictItems)
		'Wscript.Echo dictItems(dictIter)
		strRegPath = dictItems(dictIter)
		strCommand = "REG EXPORT """ & strRegPath & """ """ & backupFolderPath & dictIter & ".reg"" /y"
		'Wscript.Echo strCommand
		objBkShell.Run strCommand, 0, True
		If objBkFSO.FileExists(backupFolderPath & dictIter & ".reg") = True Then			
			Set objInputFile = objBkFSO.OpenTextFile(backupFolderPath & dictIter & ".reg", intForReading, False, intUnicode)
			If Not objInputFile.AtEndOfStream Then
				  objInputFile.SkipLine
				  objBkRegFile.Write objInputFile.ReadAll
			End If
			objInputFile.Close
			Set objInputFile = Nothing
			objBkFSO.DeleteFile backupFolderPath & dictIter & ".reg", True
		End If
	Next

	objBkRegFile.Close
	Set objBkRegFile = Nothing
End Sub

'~~~~~~ recurse through the registry keys and save Roaming keys in a dictionary ~~~~~~
Sub EnumerateRoaming(hive, key)
	Dim arrSubKeys, subkey
	if InStr(1,key,"_AD",vbTextCompare)>0 then
		'WScript.Echo key
		dictBackupRoaming.Add hiveCounter, "HKCU\"&key
		hiveCounter = hiveCounter + 1
		Exit Sub
	end if
	
	regRoaming.EnumKey hive, key, arrSubKeys
	If Not IsNull(arrSubKeys) Then
		For Each subkey In arrSubKeys
			EnumerateRoaming hive, key & "\" & subkey
		Next
	End If
End Sub

'~~~~~~ find app executable location ~~~~~~
Sub FindFileEXEPath(fFolder, fName)
    Dim objFolder, colFiles, objFile, Subfolder 'fso
	'use global fso
	'Set fso = CreateObject("Scripting.FileSystemObject")
	if dictAppsExePath.Exists(fName)=True then
		Exit Sub
	end if
	
	Set objFolder = fso.GetFolder(fFolder.Path)
    Set colFiles = objFolder.Files
    For Each objFile in colFiles
        If UCase(objFile.name) = UCase(fName) Then
            'Wscript.Echo objFile.Name
			'uses global variable
			dictAppsExePath.Add fName, objFile.Path
			Exit For
        End If
    Next

    For Each Subfolder in fFolder.SubFolders
		if InStr(1,Subfolder.Path,"Program",vbTextCompare)>0 then
			'Wscript.Echo Subfolder.Path
			FindFileEXEPath Subfolder, fName
		end if        
    Next
End Sub