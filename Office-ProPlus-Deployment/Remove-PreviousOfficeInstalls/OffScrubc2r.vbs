'*******************************************************************************
' Name: OffScrubC2R.vbs
' Author: Microsoft Customer Support Services
' Copyright (c) 2014 - 2016 Microsoft Corporation
' Script to remove Office Click To Run (C2R) products
' when a regular uninstall is no longer possible
'
' Scope: Office 2013, 2016 and O365 C2R products
'*******************************************************************************

Option Explicit

'-------------------------------------------------------------------------------
'
'   Declaration of constants
'-------------------------------------------------------------------------------


Const SCRIPTVERSION   = "2.12"
Const SCRIPTFILE      = "OffScrubC2R.vbs"
Const SCRIPTNAME      = "OffScrubC2R"
Const RETVALFILE      = "ScrubRetValFile.txt"
Const ONAME           = "Office C2R / O365"
Const HKCR            = &H80000000
Const HKCU            = &H80000001
Const HKLM            = &H80000002
Const HKU             = &H80000003
Const PRODLEN         = 13
Const SQUISHED        = 20
Const COMPRESSED      = 32
Const REG_ARP         = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
Const VB_YES          = 6
Const VB_NO           = 7

Const ERROR_SUCCESS                 = 0   'Bit #1.  0 indicates Success. Script completed successfully
Const ERROR_FAIL                    = 1   'Bit #1.  Failure bit. Indicates an overall script failure.
                                          'RESERVED bit! Returned when process is killed from task manager
Const ERROR_REBOOT_REQUIRED         = 2   'Bit #2.  Reboot bit. If set a reboot is required
Const ERROR_USERCANCEL              = 4   'Bit #3.  User Cancel bit. Controlled cancel from script UI
Const ERROR_STAGE1                  = 8   'Bit #4.  Informational. Msiexec based install was not possible
Const ERROR_STAGE2                  = 16  'Bit #5.  Critical. Not all of the intended cleanup operations could be applied
Const ERROR_INCOMPLETE              = 32  'Bit #6.  Pending file renames (del on reboot) - OR - Removal needs to run again after a system reboot.
Const ERROR_DCAF_FAILURE            = 64  'Bit #7.  Critical. Da capo al fine (second attempt) still failed.
Const ERROR_ELEVATION_USERDECLINED  = 128 'Bit #8.  Critical script error. User declined to allow mandatory script elevation
Const ERROR_ELEVATION               = 256 'Bit #9.  Critical script error. The attempt to elevate the process did not succeed
Const ERROR_SCRIPTINIT              = 512 'Bit #10. Critical script error. Initialization failed
Const ERROR_RELAUNCH                = 1024'Bit #11. Critical script error. This is a temporary value and must not be the final return code
Const ERROR_UNKNOWN                 = 2048'Bit #12 Critical script error. Script did not complete in a well defined state
Const ERROR_ALL                     = 4095'Full BitMask
Const ERROR_USER_ABORT              = &HC000013A 'RESERVED. Dec -1073741510. Critical error. Returned when user aborts with <Ctrl>+<Break> or closes the cmd window
Const ERROR_SUCCESS_CONFIG_COMPLETE = 1728
Const ERROR_SUCCESS_REBOOT_REQUIRED = 3010

'-------------------------------------------------------------------------------
'
'   Declaration of variables
'-------------------------------------------------------------------------------

Dim oFso, oMsi, oReg, oWShell, oWmiLocal, oShellApp
Dim ComputerItem, Key, Item, LogStream, TmpKey
Dim arrVersion
Dim dicKeepLis, dicApps, dicKeepFolder, dicDelRegKey, dicKeepReg
Dim dicInstalledSku, dicRemoveSku, dicKeepSku, dicC2RSuite, dicDelInUse
Dim dicDelFolder
Dim sAppData, sScrubDir, sProgramFiles, sProgramFilesX86, sCommonProgramFiles
Dim sAllusersProfile, sOSVersion, sWinDir, sWICacheDir, sCommonProgramFilesX86
Dim sProgramData, sPackageFolder, sLocalAppData, sOInstallRoot, sSkuRemoveList
Dim sOSinfo, sDefault, sTemp, sTmp, sCmd, sLogDir, sProfilesDirectory
Dim sRetVal, sScriptDir, sPackageGuid, sValue, sActiveConfiguration, sNotepad
Dim iVersionNT, iError, iProcCloseCnt
Dim f64, fLogInitialized, fNoCancel, fRemoveOse, fDetectOnly, fQuiet, fForce
Dim fC2R, fRemoveAll, fRebootRequired, fRerun, fSetRunOnce, fTestRerun
Dim fIsElevated, fNoElevate, fUserConsent, fCScript, fReturnErrorOrSuccess
Dim fClearTaskBand, fSkipSD 

'-------------------------------------------------------------------------------
'                                   Main
'
'                           Main section of script
'-------------------------------------------------------------------------------


' initialize required settings and objects
' ----------------------------------------
Initialize

' call the command line parser
'-----------------------------
ParseCmdLine

                                '-----------------------------
                                ' Stage # 0 - Basic detection |
                                '-----------------------------

LogH "Stage # 0 " & chr(34) & "Basic detection" & chr(34) 

' ensure integrity of WI metadata which could fail used APIs otherwise
'---------------------------------------------------------------------
LogH1 "Ensure Windows Installer metadata integrity " & " (" & Time & ")"
EnsureValidWIMetadata HKCU,"Software\Classes\Installer\Products", COMPRESSED
EnsureValidWIMetadata HKCR,"Installer\Products", COMPRESSED
EnsureValidWIMetadata HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products", COMPRESSED
EnsureValidWIMetadata HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components", COMPRESSED
EnsureValidWIMetadata HKCR,"Installer\Components", COMPRESSED

' build a list with installed/registered Office products
'-------------------------------------------------------
FindInstalledOProducts
If dicC2RSuite.Count > 0 Then 
    Log "Registered ARP product(s) found:"
    For Each Key In dicC2RSuite.Keys
    	Log " - " & Key & " - " & dicC2RSuite.Item(Key)
    Next 'Key
'    For Each Item in dicC2RSuite.Items
'        Log " - " & Item
'    Next 'Item
Else
    Log "No registered product(s) found"
End If

' locate the C2R %PackageFolder% and the PackageGuid
'---------------------------------------------------
sPackageFolder = ""
If RegReadValue(HKLM, "SOFTWARE\Microsoft\Office\15.0\ClickToRun", "PackageFolder", sValue, "REG_SZ") Then
	sPackageFolder = sValue
ElseIf RegReadValue(HKLM, "SOFTWARE\Microsoft\Office\16.0\ClickToRun", "PackageFolder", sPackageFolder, "REG_SZ") Then
	sPackageFolder = sValue
ElseIf RegReadValue(HKLM, "SOFTWARE\Microsoft\Office\ClickToRun", "PackageFolder", sPackageFolder, "REG_SZ") Then
	sPackageFolder = sValue
End If
' if sPackageFolder is invalid set it to the c2r registry reference string
If NOT Len(sPackageFolder) > 0 OR IsNull(sPackageFolder) Then 
	If oFso.FolderExists(oWShell.ExpandEnvironmentStrings("%programfiles%") & "\Microsoft Office 15") Then
		sPackageFolder = oWShell.ExpandEnvironmentStrings("%programfiles%") & "\Microsoft Office 15"
	ElseIf oFso.FolderExists(oWShell.ExpandEnvironmentStrings("%programfiles%") & "\Microsoft Office 16") Then
		sPackageFolder = oWShell.ExpandEnvironmentStrings("%programfiles%") & "\Microsoft Office 16"
	ElseIf oFso.FolderExists(oWShell.ExpandEnvironmentStrings("%programfiles%") & "\Microsoft Office\PackageManifests") Then
		sPackageFolder = oWShell.ExpandEnvironmentStrings("%programfiles%") & "\Microsoft Office"
	ElseIf oFso.FolderExists(oWShell.ExpandEnvironmentStrings("%programfiles(x86)%") & "\Microsoft Office\PackageManifests") Then
		sPackageFolder = oWShell.ExpandEnvironmentStrings("%programfiles(x86)%") & "\Microsoft Office"
	End If
End If

sPackageGuid = ""
If RegReadValue(HKLM, "SOFTWARE\Microsoft\Office\15.0\ClickToRun", "PackageGUID", sValue, "REG_SZ") Then
	sPackageGuid = sValue
ElseIf RegReadValue(HKLM, "SOFTWARE\Microsoft\Office\16.0\ClickToRun", "PackageGUID", sValue, "REG_SZ") Then
	sPackageGuid = sValue
ElseIf RegReadValue(HKLM, "SOFTWARE\Microsoft\Office\ClickToRun", "PackageGUID", sValue, "REG_SZ") Then
	sPackageGuid = sValue
End If

' Init complete. Reset the return value
'--------------------------------------
ClearError ERROR_SCRIPTINIT


                                '-----------------------
                                ' Stage # 1 - Uninstall |
                                '-----------------------

LogH "Stage # 1 " & chr(34) & "Uninstall" & chr(34)

' clean O15 SPP
'--------------
LogH1 "Clean OSPP" 
CleanOSPP

' end all running Office applications
'------------------------------------
LogH1 "End running processes" 
If NOT dicKeepSku.Count > 0 Then ClearShellIntegrationReg
CloseOfficeApps

' remove scheduled tasks which might interfere with uninstall
'------------------------------------------------------------
DelSchtasks

' unpin shortcuts
'----------------
' need to unpin as long as the shortcuts are still valid!
LogH1 "Clean shortcuts" 
CleanShortcuts sAllusersProfile, True, True
CleanShortcuts sProfilesDirectory, True, True

' uninstall
'----------
LogH1 "Remove " & ONAME  
Uninstall

                                '---------------------
                                ' Stage # 2 - CleanUp |
                                '---------------------
LogH "Stage # 2 " & chr(34) & "CleanUp" & chr(34)

' Cleanup registry data
'----------------------
RegWipe

' Cleanup files
'--------------
FileWipe

' for test purposes only!
If fTestRerun Then
    LogH2 "Enforcing 'Rerun' mode for test purposes"
    fRebootRequired = True
    SetError ERROR_REBOOT_REQUIRED
    Rerun
End If

' Ensure Explorer runs
RestoreExplorer

' Exit
ExitScript

                                    '------------------
                                    ' Stage # 3 - Exit |
                                    '------------------

'-------------------------------------------------------------------------------
'   ExitScript
'
'   Returncode and reboot handler 
'-------------------------------------------------------------------------------
Sub ExitScript
    Dim sPrompt

    ' Update cached error and quit
    '-----------------------------
    If NOT CBool(iError AND (ERROR_FAIL + ERROR_INCOMPLETE)) Then RegDeleteValue HKCU, "SOFTWARE\Microsoft\Office\15.0\CleanC2R", "Rerun", False
    SetRetVal iError

    ' log result
    If CBool(iError AND ERROR_INCOMPLETE) Then 
        LogH2 "Removal result: " & iError & " - INCOMPLETE. Uninstall requires a system reboot to complete."
    Else
        sTmp = " - SUCCESS"
        If CBool(iError AND ERROR_USERCANCEL) Then sTmp = " - USER CANCELED"
        If CBool(iError AND ERROR_FAIL) Then sTmp = " - FAIL"
        LogH2 "Removal result: " & iError & sTmp
    End If
    If CBool(iError AND ERROR_FAIL) Then
        If CBool(iError AND ERROR_REBOOT_REQUIRED) Then Log " - Reboot required"
        If CBool(iError AND ERROR_USERCANCEL) Then Log " - User cancel"
        If CBool(iError AND ERROR_STAGE1) Then Log " - Msiexec failed"
        If CBool(iError AND ERROR_STAGE2) Then Log " - Cleanup failed"
        If CBool(iError AND ERROR_INCOMPLETE) Then Log " - Removal incomplete. Rerun after reboot needed"
        If CBool(iError AND ERROR_DCAF_FAILURE) Then Log " - Second attempt cleanup still incomplete"
        If CBool(iError AND ERROR_ELEVATION_USERDECLINED) Then Log " - User declined elevation"
        If CBool(iError AND ERROR_ELEVATION) Then Log " - Elevation failed"
        If CBool(iError AND ERROR_SCRIPTINIT) Then Log " - Initialization error"
        If CBool(iError AND ERROR_RELAUNCH) Then Log " - Unhandled error during relaunch attempt"
        If CBool(iError AND ERROR_UNKNOWN) Then Log " - Unknown error"
        ' ERROR_USER_ABORT is only valid for the temporary cached error file
        'If CBool(iError AND ERROR_USER_ABORT) Then Log " - Process terminated by user"
    End If

    LogH2 "Removal end." 

    ' Check if we need to show a simplified return code
    ' 0 = Success
    ' Non Zero = Error
     If CBool(iError AND ERROR_FAIL) AND fReturnErrorOrSuccess Then
        Dim fOverallSuccess
        fOverallSuccess = True
        If CBool(iError AND ERROR_USERCANCEL) Then fOverallSuccess = False
        If CBool(iError AND ERROR_STAGE2) Then fOverallSuccess = False
        If CBool(iError AND ERROR_DCAF_FAILURE) Then fOverallSuccess = False
        If CBool(iError AND ERROR_ELEVATION_USERDECLINED) Then fOverallSuccess = False
        If CBool(iError AND ERROR_ELEVATION) Then fOverallSuccess = False
        If CBool(iError AND ERROR_SCRIPTINIT) Then fOverallSuccess = False
        If CBool(iError AND ERROR_RELAUNCH) Then fOverallSuccess = False
        If CBool(iError AND ERROR_UNKNOWN) Then fOverallSuccess = False

        If fOverallSuccess Then iError = ERROR_SUCCESS

        sTmp = "ReturnErrorOrSuccess switch has been set. The current value return code translates to: "
        If fOverallSuccess Then 
            iError = ERROR_SUCCESS
            Log sTmp & "SUCCESS"
        Else
            Log sTmp & "ERROR"
        End If

    End If
   
    ' Reboot handling
    If fRebootRequired Then
        sPrompt = "In order to complete uninstall, a system reboot is necessary. Would you like to reboot now?"
        If NOT fQuiet Then
            If MsgBox(sPrompt, vbYesNo, SCRIPTNAME & " - Reboot Required") = VB_YES Then
                Dim colOS, oOS
                Dim oWmiReboot
                Set oWmiReboot = GetObject("winmgmts:{impersonationLevel=impersonate,(Shutdown)}!\\.\root\cimv2")
                Set colOS = oWmiReboot.ExecQuery ("Select * from Win32_OperatingSystem")
                For Each oOS in colOS
                    oOS.Reboot()
                Next
            End If
        End If
    End If

    wscript.quit iError
End Sub 'ExitScript

'-------------------------------------------------------------------------------
'                                  End  Main
'
'                           End of Main section
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'   Initialize
'
'   Configure defaults and initialize all required objects
'-------------------------------------------------------------------------------
Sub Initialize ()
    Dim iCnt

    ' set defaults
    '-------------
    iError = ERROR_SUCCESS
    iProcCloseCnt = 0
    sLogDir = ""
    sPackageFolder = ""
    f64 = False
    fCScript = False
    fLogInitialized = False
    fNoCancel = False
    fRemoveOse = False
    fDetectOnly = False
    fQuiet = False
    fForce = False
    fC2R = True
    fRebootRequired = False
    fRerun = False
    fTestRerun = False
    fIsElevated = False
    fNoElevate = False
    fSetRunOnce = False
    fUserConsent = False
    fReturnErrorOrSuccess = False
    fSkipSD = False
    fClearTaskBand = False

    ' create required objects
    '------------------------
    Set oWmiLocal   = GetObject("winmgmts:\\.\root\cimv2")
    Set oWShell     = CreateObject("Wscript.Shell")
    Set oShellApp   = CreateObject("Shell.Application")
    Set oFso        = CreateObject("Scripting.FileSystemObject")
    Set oMsi        = CreateObject("WindowsInstaller.Installer")
    Set oReg        = GetObject("winmgmts:\\.\root\default:StdRegProv")

    ' get environment path values
    '----------------------------
    sAppData            = oWShell.ExpandEnvironmentStrings("%appdata%")
    sLocalAppData       = oWShell.ExpandEnvironmentStrings("%localappdata%")
    sTemp               = oWShell.ExpandEnvironmentStrings("%temp%")
    sAllUsersProfile    = oWShell.ExpandEnvironmentStrings("%allusersprofile%")
    RegReadValue HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList", "ProfilesDirectory", sProfilesDirectory, "REG_EXPAND_SZ"
    If NOT oFso.FolderExists(sProfilesDirectory) Then 
        sProfilesDirectory  = oFso.GetParentFolderName(oWShell.ExpandEnvironmentStrings("%userprofile%"))
    End If
    sProgramFiles       = oWShell.ExpandEnvironmentStrings("%programfiles%")
    'sProgramFilesX86   = deferred. Depends on operating system architecture check
    sCommonProgramFiles = oWShell.ExpandEnvironmentStrings("%commonprogramfiles%")
    'sCommonProgramFilesX86 = deferred. Depends on operating system architecture check
    sProgramData        = oWSHell.ExpandEnvironmentStrings("%programdata%")
    sWinDir             = oWShell.ExpandEnvironmentStrings("%windir%")
    'sPackageFolder      = deferred
    sWICacheDir         = sWinDir & "\" & "Installer"
    sScrubDir           = sTemp & "\" & SCRIPTNAME
    sScriptDir          = wscript.ScriptFullName
    sScriptDir          = Left(sScriptDir, InStrRev(sScriptDir, "\"))
    sNotepad            = sWinDir & "\notepad.exe"

    ' ensure 64 bit host if needed
    If InStr(LCase(wscript.path), "syswow64") > 0 Then RelaunchAs64Host

    ' create the temp folder
    '-----------------------
    If Not oFso.FolderExists(sScrubDir) Then oFso.CreateFolder sScrubDir

    ' set the default logging directory
    '----------------------------------
    sLogDir = sScrubDir

    ' detect bitness of the operating system
    '----------------------------------------
    Set ComputerItem = oWmiLocal.ExecQuery("Select * from Win32_ComputerSystem")
    For Each Item In ComputerItem
        f64 = Instr(Left(Item.SystemType, 3), "64") > 0
    Next
    If f64 Then sProgramFilesX86 = oWShell.ExpandEnvironmentStrings("%programfiles(x86)%")
    If f64 Then sCommonProgramFilesX86 = oWShell.ExpandEnvironmentStrings("%CommonProgramFiles(x86)%")

    ' update error flag
    '------------------
    SetError ERROR_SCRIPTINIT

    ' get Win32_OperatingSystem details
    '----------------------------------
    Set ComputerItem = oWmiLocal.ExecQuery("Select * from Win32_OperatingSystem")
    For Each Item in ComputerItem 
        sOSinfo = sOSinfo & Item.Caption 
        sOSinfo = sOSinfo & Item.OtherTypeDescription
        sOSinfo = sOSinfo & ", " & "SP " & Item.ServicePackMajorVersion
        sOSinfo = sOSinfo & ", " & "Version: " & Item.Version
        sOsVersion = Item.Version
        sOSinfo = sOSinfo & ", " & "Codepage: " & Item.CodeSet
        sOSinfo = sOSinfo & ", " & "Country Code: " & Item.CountryCode
        sOSinfo = sOSinfo & ", " & "Language: " & Item.OSLanguage
    Next

    ' get VersionNT number
    '---------------------
    arrVersion = Split(sOsVersion, Delimiter(sOsVersion))
    iVersionNt = CInt(arrVersion(0)) * 100 + CInt(arrVersion(1))

    ' ensure sufficient registry permisions
    '--------------------------------------
    fIsElevated = CheckRegPermissions
    If NOT fIsElevated AND NOT fNoElevate Then
        ' try to relaunch elevated
        RelaunchElevated

        ' can't relaunch. Exit out
        SetError ERROR_ELEVATION
        If UCase(Mid(Wscript.FullName, Len(Wscript.Path) + 2, 1)) = "C" Then
            If Not fLogInitialized Then CreateLog
            Log "Error: Insufficient registry access permissions - exiting"
        End If
        SetRetVal iError
        'wscript.quit iError
        ExitScript
    End If

    ' clear error flags
    '------------------
    ClearError ERROR_ELEVATION
    ClearError ERROR_SCRIPTINIT

    ' ensure CScript as engine
    '------------------------
    fCScript = UCase(Mid(Wscript.FullName, Len(Wscript.Path) + 2, 1)) = "C"
    If NOT fCScript AND NOT fQuiet Then RelaunchAsCScript

    ' set retval for file based logic
    '--------------------------------
    ' value needs to be kept on 'user abort'
    SetRetVal ERROR_USER_ABORT

    ' create dictionary objects
    '--------------------------
    Set dicInstalledSku = CreateObject("Scripting.Dictionary")
    Set dicRemoveSku = CreateObject("Scripting.Dictionary")
    Set dicKeepSku = CreateObject("Scripting.Dictionary")
    Set dicKeepLis = CreateObject("Scripting.Dictionary")
    Set dicKeepFolder = CreateObject("Scripting.Dictionary")
    Set dicApps = CreateObject("Scripting.Dictionary")
    Set dicDelRegKey = CreateObject("Scripting.Dictionary")
    Set dicKeepReg = CreateObject("Scripting.Dictionary")
    Set dicC2RSuite = CreateObject("Scripting.Dictionary")
    Set dicDelInUse = CreateObject("Scripting.Dictionary")
    Set dicDelFolder = CreateObject("Scripting.Dictionary")

    ' add initial known .exe files that need to be closed
    '----------------------------------------------------
    dicApps.Add "appvshnotify.exe", "appvshnotify.exe"
    dicApps.Add "integratedoffice.exe", "integratedoffice.exe"
    dicApps.Add "integrator.exe", "integrator.exe"
    dicApps.Add "firstrun.exe", "firstrun.exe"
    'Adding setup.exe to the hard list of processes that are shut down will potentially break wrappers that invoke OffScrub
    'dicApps.Add "setup.exe", "setup.exe"
    dicApps.Add "communicator.exe", "communicator.exe"
    dicApps.Add "msosync.exe", "msosync.exe"
    dicApps.Add "OneNoteM.exe", "OneNoteM.exe"
    dicApps.Add "iexplore.exe", "iexplore.exe"
    dicApps.Add "mavinject32.exe", "mavinject32.exe"
    dicApps.Add "werfault.exe", "werfault.exe"
    dicApps.Add "perfboost.exe", "perfboost.exe"
    dicApps.Add "roamingoffice.exe", "roamingoffice.exe"
    ' SP1 additions / changes
    dicApps.Add "officeclicktorun.exe", "officeclicktorun.exe"
    dicApps.Add "officeondemand.exe", "officeondemand.exe"
    dicApps.Add "OfficeC2RClient.exe", "OfficeC2RClient.exe"

End Sub 'Initialize

'-------------------------------------------------------------------------------
'   ParseCmdLine
'
'   Command line parser
'-------------------------------------------------------------------------------
Sub ParseCmdLine

    Dim iCnt, iArgCnt
    Dim arrArguments, sArguments
    Dim sArg0
    
    iArgCnt = Wscript.Arguments.Count
    If iArgCnt > 0 Then
        If wscript.Arguments(0) = "UAC" Then
            If wscript.arguments.count = 1 Then iArgCnt = 0
        End If
    End If
    If iArgCnt = 0 Then
        Select Case UCase(wscript.ScriptName)
        Case Else
            'Create the log
            CreateLog
            FindInstalledOProducts
            sDefault = "ALL"
            arrArguments = Split(Trim(sDefault), " ")
            If UBound(arrArguments) = -1 Then ReDim arrArguments(0)
        End Select
    Else
        ReDim arrArguments(iArgCnt-1)
        For iCnt = 0 To (iArgCnt-1)
            arrArguments(iCnt) = UCase(Wscript.Arguments(iCnt))
            sArguments = sArguments & arrArguments(iCnt) & " "
        Next 'iCnt
    End If 'iArgCnt = 0

    ' hardcode to full removal
    sArg0 = "ALL"

    Select Case UCase(sArg0)
    Case "?"
        ShowSyntax
    Case "ALL"
        fRemoveAll = True
        fRemoveOse = False
    Case "C2R"
        fC2R = True
        fRemoveAll = False
        fRemoveOse = False
    Case Else
        fRemoveAll = False
        fRemoveOse = False
        sSkuRemoveList = sArg0
    End Select
    
    For iCnt = 0 To UBound(arrArguments)
        Select Case arrArguments(iCnt)
        Case "?", "/?", "-?"
            ShowSyntax
        
        Case "/L", "/LOG"
            fLogInitialized = False
            If UBound(arrArguments) > iCnt Then
                If oFso.FolderExists(arrArguments(iCnt + 1)) Then 
                    sLogDir = arrArguments(iCnt + 1)
                Else
                    On Error Resume Next
                    oFso.CreateFolder(arrArguments(iCnt + 1))
                    If Err <> 0 Then sLogDir = sScrubDir Else sLogDir = arrArguments(iCnt + 1)
                End If
            End If
        
        Case "/N", "/NOCANCEL"
            fNoCancel = True
        
        Case "/NE", "/NOELEVATE"
            fNoElevate = True

        Case "/O", "/OSE"
            fRemoveOse = True
        
        Case "/Q", "/QUIET"
            fQuiet = True
        
        Case "/RETERRORSUCCESS", "/RETURNERRORORSUCCESS", "/REOS"
            fReturnErrorOrSuccess = True

        Case "/S", "/SKIPSD", "/SKIPSHORTCUTDETECTION"
            fSkipSD = True
        
        ' for test purposes only!
        Case "/TR", "/TESTRERUN"
            fTestRerun = True
        Case Else
        End Select
    Next 'iCnt
    If Not fLogInitialized Then CreateLog
    LogH2 "Arguments: " & sArguments & vbCrLf
    
End Sub 'ParseCmdLine

'-------------------------------------------------------------------------------
'   ShowSyntax
'
'   Show the expected syntax for the script usage
'-------------------------------------------------------------------------------
Sub ShowSyntax
    Wscript.Echo vbCrLf & _
             SCRIPTFILE & " V " & SCRIPTVERSION & vbCrLf & _
             "Copyright (c) Microsoft Corporation. All Rights Reserved" & vbCrLf & vbCrLf & _
             SCRIPTFILE & " - Remove " & ONAME & vbCrLf & _
             "when a regular uninstall is no longer possible" & vbCrLf & vbCrLf & _
             "Usage:" & vbTab & SCRIPTFILE & vbCrLf & vbCrLf & _
	     vbTab & "ALL                         ' Remove all products and features" & vbCrLf & _
             vbTab & "/?                          ' Displays this help" & vbCrLf & _
             vbTab & "/Log [LogfolderPath]        ' Custom folder for log files" & vbCrLf & _
             vbTab & "/SkipSD                     ' Skips the ShortcutDetection in local profiles" & vbCrLf & _
             vbTab & "/NoCancel                   ' Setup.exe and Msiexec.exe have no Cancel button" & vbCrLf & _
             vbTab & "/Quiet                      ' Script, Setup.exe and Msiexec.exe run quiet with no UI" & vbCrLf & _
             vbTab & "/ReturnErrorOrSuccess       ' Returns 0 for a successful removal. Non-Zero if not." & vbCrLf & _
	     vbTab & "/NE                         ' No-elevation of uninstall process" & vbCrLf & _
	     vbTab & "/OSE                        ' Remove OSE" & vbCrLf & _
	     vbTab & "/TR                         ' Test re-run" & vbCrLf
    Wscript.Quit
End Sub 'ShowSyntax

'-------------------------------------------------------------------------------
'   FindInstalledOProducts
'
'   Office configuration products are listed with their configuration product
'   name in the "Uninstall" key.
'-------------------------------------------------------------------------------
Sub FindInstalledOProducts
    Dim ArpItem, prod, cult
    Dim sCurKey, sValue, sConfigName, sCulture, sDisplayVersion, sVersionFallback
    Dim sUninstallString, sProd
    Dim iLeft, iRight
    Dim arrKeys, arrProducts, arrCultures
    Dim fSystemComponent0, fDisplayVersion, fUninstallString

	Const REG_ARP                 = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
	Const REG_O15RPROPERTYBAG     = "SOFTWARE\Microsoft\Office\15.0\ClickToRun\propertyBag\"
	Const REG_O15C2RCONFIGURATION = "SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration\"
	Const REG_O15C2RPRODUCTIDS    = "SOFTWARE\Microsoft\Office\15.0\ClickToRun\ProductReleaseIDs\Active\"
	Const REG_O16C2RCONFIGURATION = "SOFTWARE\Microsoft\Office\16.0\ClickToRun\Configuration\"
	Const REG_O16C2RPRODUCTIDS    = "SOFTWARE\Microsoft\Office\16.0\ClickToRun\ProductReleaseIDs\Active\"
	Const REG_C2RCONFIGURATION    = "SOFTWARE\Microsoft\Office\ClickToRun\Configuration\"
	Const REG_C2RPRODUCTIDS       = "SOFTWARE\Microsoft\Office\ClickToRun\ProductReleaseIDs\"


    If dicInstalledSku.Count > 0 Then Exit Sub 'Already done from command line parser
    
    fDisplayVersion = False

    ' identify C2R products
    LogH1 "Detect installed products "
    
	LogOnly "Check for O15 C2R products"
	' Check O15 Configuration key
    If RegReadValue(HKLM, REG_O15C2RCONFIGURATION, "ProductReleaseIds", sValue, "REG_SZ") Then
        arrProducts = Split(sValue, ",")
        fDisplayVersion = RegReadValue(HKLM, REG_O15C2RPRODUCTIDS & "culture", "x-none", sVersionFallback, "REG_SZ")
        If NOT Err = 0 Then
            Err.Clear
        Else
            ' get version from active with fallback on configuration
            For Each prod in arrProducts
                LogOnly "Found O15 C2R product in Configuration: " & prod
                ' update product dictionary
                If NOT dicInstalledSku.Exists(LCase(prod)) Then
                	LogOnly "add new product to dictionary: " & LCase(prod)
                	dicInstalledSku.Add LCase(prod), sVersionFallback
                End If
            Next 'prod
        End If
    End If
    
    ' Check O15 PropertyBag key
    If RegReadValue(HKLM, REG_O15RPROPERTYBAG, "productreleaseid", sValue, "REG_SZ") Then
        arrProducts = Split(sValue, ",")
        fDisplayVersion = RegReadValue(HKLM, REG_O15C2RPRODUCTIDS & "culture", "x-none", sVersionFallback, "REG_SZ")
        If NOT Err = 0 Then
            Err.Clear
        Else
            For Each prod in arrProducts
                LogOnly "Found O15 C2R product in PropertyBag: " & prod
                ' update product dictionary
                If NOT dicInstalledSku.Exists(LCase(prod)) Then
                	LogOnly "add new product to dictionary: " & LCase(prod)
                	dicInstalledSku.Add LCase(prod), sVersionFallback
                End If
            Next 'prod
        End If
    End If
    
	'O16 section
	LogOnly "Check for Office C2R products (>=QR8)"
	' Check Office Configuration key
    If RegReadValue(HKLM, REG_C2RPRODUCTIDS, "ActiveConfiguration", sActiveConfiguration, "REG_SZ") Then
    	' Get DisplayVersion
    	'Try QR8 logic first
		fDisplayVersion = RegReadValue(HKLM, REG_C2RPRODUCTIDS & sActiveConfiguration & "\culture", "x-none", sVersionFallback, "REG_SZ")
    	If RegEnumKey(HKLM, REG_C2RPRODUCTIDS & sActiveConfiguration & "\culture", arrCultures) Then
    		For Each cult In arrCultures
    			If InStr(LCase(cult), "x-none") > 0 Then
    				fDisplayVersion = RegReadValue(HKLM, REG_C2RPRODUCTIDS & sActiveConfiguration & "\culture\" & cult, "Version", sVersionFallback, "REG_SZ")
    			End If
    		Next 'cult
    	End If 
    	' Update product dic
    	If RegEnumKey(HKLM, REG_C2RPRODUCTIDS & sActiveConfiguration, arrProducts) Then
    		For Each prod In arrProducts
    			sProd = LCase(prod)
    			If InStr(sProd, ".") > 0 Then sProd = Left(sProd, InStr(sProd, ".") - 1)
    			Select Case LCase(sProd)
    			Case "culture", "stream"
    			Case Else
	                LogOnly "Found Office C2R product in Configuration: " & prod
	                ' update product dictionary
	                If NOT dicInstalledSku.Exists(sProd) Then
		                LogOnly "add new product to dictionary: " & sProd
		                If RegReadValue(HKLM, REG_C2RPRODUCTIDS & sActiveConfiguration & "\" & prod & "\x-none", "Version", sDisplayVersion, "REG_SZ") Then
		                	dicInstalledSku.Add sProd, sDisplayVersion
		                Else
	                		dicInstalledSku.Add sProd, sVersionFallback
	                	End If
	                End If
    			End Select
    		Next 'prod
    	End If 'arrProducts
    End If 'ActiveConfiguration

	LogOnly "Check for Office C2R products (QR7)"
	' Check Office Configuration key
	If RegReadValue(HKLM, REG_C2RCONFIGURATION, "ProductReleaseIds", sValue, "REG_SZ") Then
	    arrProducts = Split(sValue, ",")
	    If Not fDisplayVersion Then fDisplayVersion = RegReadValue(HKLM, REG_C2RPRODUCTIDS & "Active\culture", "x-none", sVersionFallback, "REG_SZ")
	    If NOT Err = 0 Then
	        Err.Clear
	    Else
	        For Each prod in arrProducts
	            LogOnly "Found Office C2R product in Configuration: " & prod
	            ' update version tracking
	            If NOT dicInstalledSku.Exists(LCase(prod)) Then
	            	LogOnly "add new product to dictionary: " & LCase(prod)
	            	dicInstalledSku.Add LCase(prod), sVersionFallback
	            End If
	        Next 'prod
	    End If
	End If

	LogOnly "Check for O16 C2R products (QR6)"
	' Check O16 Configuration key
    If RegReadValue(HKLM, REG_O16C2RCONFIGURATION, "ProductReleaseIds", sValue, "REG_SZ") Then
        arrProducts = Split(sValue, ",")
        If Not fDisplayVersion Then fDisplayVersion = RegReadValue(HKLM, REG_O16C2RPRODUCTIDS & "culture", "x-none", sVersionFallback, "REG_SZ")
        If NOT Err = 0 Then
            Err.Clear
        Else
            For Each prod in arrProducts
                LogOnly "Found O16 (QR6) C2R product in Configuration: " & prod
                ' update product dictionary
                If NOT dicInstalledSku.Exists(LCase(prod)) Then
                	LogOnly "add new product to dictionary: " & prod
                	dicInstalledSku.Add LCase(prod), sVersionFallback
                End If
            Next 'prod
        End If
    End If

    LogOnly "Check ARP for Office C2R products"
    ' ARP
    RegEnumKey HKLM, REG_ARP, arrKeys
    If IsArray(arrKeys) Then
        For Each ArpItem in arrKeys
            ' filter on Office C2R products
            sCurKey = REG_ARP & ArpItem & "\"
            fUninstallString = RegReadValue(HKLM, sCurKey, "UninstallString", sValue, "REG_SZ")
            If (fUninstallString And( (InStr(UCase(sValue), UCase("Microsoft Office 1")) > 0) Or (InStr(UCase(sValue), UCase("OfficeClickToRun.exe")) > 0) )) Then
                'get Version
                fDisplayVersion = RegReadValue(HKLM, sCurKey, "DisplayVersion", sDisplayVersion, "REG_SZ")
                'extract the productreleaseid
                sValue = Trim(sValue)
                prod = Trim(Mid(sValue, InStrRev(sValue, " ")))
                prod = Replace(prod, "productstoremove=", "")
                If InStr(prod, "_") > 0 Then
                    prod = Left(prod, InStr(prod, "_") - 1)
                End If
                If InStr(prod, ".1") > 0 Then
                    prod = Left(prod, InStr(prod, ".1") - 1)
                End If
                LogOnly "Found C2R product in ARP: " & prod
                If NOT dicInstalledSku.Exists(LCase(prod)) Then
                	LogOnly "add new product to dictionary: " & prod
                	dicInstalledSku.Add LCase(prod), sDisplayVersion
                End If
                ' categorize the SKU as C2R
                If NOT dicC2RSuite.Exists(ArpItem) Then dicC2RSuite.Add ArpItem, prod & " - " & sDisplayVersion 
            Else
            
	            'Legacy logic keep for compat reasons
	            sValue = ""
	            sDisplayVersion = ""
	            fSystemComponent0 = NOT (RegReadValue(HKLM, sCurKey, "SystemComponent", sValue, "REG_DWORD") AND (sValue = "1"))
	            fDisplayVersion = RegReadValue(HKLM, sCurKey, "DisplayVersion", sValue, "REG_SZ")
	            If fDisplayVersion Then
	                sDisplayVersion = sValue
	                If Len(sValue) > 1 Then
	                	On Error Resume Next
	                    fDisplayVersion = (CInt(Left(sValue, 2)) > 14)
	                    If Not Err <> 0 Then Err.Clear
	                Else
	                    fDisplayVersion = False
	                End If
	            End If
	            fUninstallString = RegReadValue(HKLM, sCurKey, "UninstallString", sUninstallString, "REG_SZ")
	            
	            ' filter on C2R configuration SKU
	            If (fUninstallString And( (InStr(UCase(sUninstallString), UCase("Microsoft Office 1")) > 0) Or (InStr(UCase(sUninstallString), UCase("OfficeClickToRun.exe")) > 0) )) Then
	                ' Extract the ProductReleaseID
	                If InStr(sUninstallString, "productstoremove=") > 0 Then
		                sConfigName = Trim(Mid(sValue, InStrRev(sUninstallString, " ")))
		                sConfigName = Replace(sConfigName, "productstoremove=", "")
		                If InStr(prod, "_") > 0 Then
		                    sConfigName = Left(sConfigName, InStr(sConfigName, "_") - 1)
		                End If
	                Else
		                iLeft = InStr(ArpItem, " - ") + 2
		                iRight = InStr(iLeft, ArpItem, " - ") - 1
		                If iRight > 0 Then
		                    sConfigName = Trim(Mid(ArpItem, iLeft, (iRight - iLeft)))
		                    sCulture = Mid(ArpItem, iRight + 3)
		                Else
		                    sConfigName = Trim(Left(ArpItem, iLeft - 3))
		                    sCulture = Mid(ArpItem, iLeft)
		                End If
		                sConfigName = Replace(sConfigName, "Microsoft", "")
		                sConfigName = Replace(sConfigName, "Office", "")
		                sConfigName = Replace(sConfigName, "Professional", "Pro")
		                sConfigName = Replace(sConfigName, "Standard", "Std")
		                sConfigName = Replace(sConfigName, "(Technical Preview)", "")
		                sConfigName = Replace(sConfigName, "15", "")
		                sConfigName = Replace(sConfigName, "16", "")
		                sConfigName = Replace(sConfigName, "2013", "")
		                sConfigName = Replace(sConfigName, "2016", "")
		                sConfigName = Replace(sConfigName, " ", "")
		                sConfigName = Replace(sConfigName, "Project", "Prj")
		                sConfigName = Replace(sConfigName, "Visio", "Vis")
	                End If
	                If NOT dicInstalledSku.Exists(LCase(sConfigName)) Then
	                	LogOnly "add new product to dictionary (ARP Legacy): " & sConfigName
	                	dicInstalledSku.Add LCase(sConfigName), sDisplayVersion
	                End If
	                ' categorize the SKU as C2R
	                If NOT dicC2RSuite.Exists(ArpItem) Then dicC2RSuite.Add ArpItem, sConfigName & " - " & sDisplayVersion 
	            ElseIf (fDisplayVersion AND (InStr(UCase(ArpItem), UCase("OFFICE15.")) > 0 Or InStr(UCase(ArpItem), UCase("OFFICE16.")) > 0)) Then
	                ' classic .msi install SKU
	                iLeft = InStr(ArpItem, ".") + 1
	                iRight = InStr(iLeft, ArpItem, "-") - 1
	                sConfigName = Mid(ArpItem, iLeft)
	                sCulture = ""
	                If NOT dicKeepSku.Exists(ArpItem) Then dicKeepSku.Add ArpItem, sConfigName & " - " & sDisplayVersion
	            End If
	            
	            ' Other products
	            If InScope(ArpItem) Then
	                Select Case Mid(ArpItem,11,4)
	                ' 007E = Licensing
	                ' 008F = Licensing
	                ' 008C = Extensibility Components
	                ' 00DD = Extensibility Components 64 bit
	                Case "007E", "008F", "008C", "00DD"
	                    sConfigName = "Habanero"
	                    RegReadValue HKLM, sCurKey, "DisplayName", sConfigName, "REG_SZ"
	                    If NOT dicInstalledSku.Exists(LCase(ArpItem)) Then
	                    	LogOnly "add new product to dictionary (ARP Integraton Components): " & ArpItem
	                    	dicInstalledSku.Add LCase(ArpItem), sDisplayVersion
	                    End If
	                    If NOT dicC2RSuite.Exists(ArpItem) Then dicC2RSuite.Add ArpItem, sConfigName & " - " & sDisplayVersion
	                Case "24E1", "237A"
	                    sConfigName = "MSOIDLOGIN"
	                    If NOT dicInstalledSku.Exists(LCase(ArpItem)) Then
	                    	LogOnly "add new product to dictionary (ARP MSOIDLogin): " & ArpItem
	                    	dicInstalledSku.Add LCase(ArpItem), sDisplayVersion
	                    End If
	                    If NOT dicC2RSuite.Exists(ArpItem) Then dicC2RSuite.Add ArpItem, sConfigName & " - " & sDisplayVersion
	                Case Else
	                    If NOT dicInstalledSku.Exists(LCase(ArpItem)) Then
	                    	LogOnly "add new product to dictionary (ARP other): " & ArpItem
	                    	dicInstalledSku.Add LCase(ArpItem), sDisplayVersion
	                    End If
	                End Select
	            Else
                    ' not in scope for c2r removal!
	            End If 'InScope  
	            ' End legacy logic
	            
            End If
        Next 'ArpItem
    End If
    
End Sub 'FindInstalledOProducts

'-------------------------------------------------------------------------------
'   EnsureValidWIMetadata
'
'   Ensures that only valid metadata entries exist to avoid API failures.
'   Invalid entries will be removed
'-------------------------------------------------------------------------------
Sub EnsureValidWIMetadata(hDefKey, sKey, iValidLength)
    Dim arrKeys
    Dim SubKey

    If Len(sKey) > 1 Then
        If Right(sKey, 1) = "\" Then sKey = Left(sKey, Len(sKey) - 1)
    End If

    If RegEnumKey(hDefKey, sKey, arrKeys) Then
        For Each SubKey in arrKeys
            If NOT Len(SubKey) = iValidLength Then
                RegDeleteKey hDefKey, sKey & "\" & SubKey & "\"
            End If
        Next 'SubKey
    End If
End Sub 'EnsureValidWIMetadata

'-------------------------------------------------------------------------------
'   CleanOSPP
'
'   Clean out licenses from the Office Software Protection Platform 
'-------------------------------------------------------------------------------
Sub CleanOSPP
    Dim oProductInstances, pi
    Dim sCleanOSPP, sCmd, sRetVal

    CONST OfficeAppId = "0ff1ce15-a989-479d-af46-f275c6370663"  'Office 2013

    sCleanOSPP = "x64\CleanOSPP.exe"
    If Not f64 Then sCleanOSPP = "x86\CleanOSPP.exe"
    If oFso.FileExists(sScriptDir & sCleanOSPP) Then
        sCmd = sScriptDir & sCleanOSPP
        Log "   Running: " & sCmd
        On Error Resume Next
        sRetVal = oWShell.Run(sCmd, 0, True)
        Log "   Return value: " & sRetVal
        On Error Goto 0
        Exit Sub
    End If
    
    On Error Resume Next
    If NOT (dicC2RSuite.Count > 0 OR dicKeepSku.Count > 0) Then
        Log "Skip CleanOSPP"
        Exit Sub
    End If
    
    ' Initialize the software protection platform object with a filter on Office 2013 products
    If iVersionNT > 601 Then
        Set oProductInstances = oWmiLocal.ExecQuery("SELECT ID, ApplicationId, PartialProductKey, Name, ProductKeyID FROM SoftwareLicensingProduct WHERE ApplicationId = '" & OfficeAppId & "' " & "AND PartialProductKey <> NULL")
    Else
        Set oProductInstances = oWmiLocal.ExecQuery("SELECT ID, ApplicationId, PartialProductKey, Name, ProductKeyID FROM OfficeSoftwareProtectionProduct WHERE ApplicationId = '" & OfficeAppId & "' " & "AND PartialProductKey <> NULL")
    End If

    ' Remove all licenses
    For Each pi in oProductInstances
        If NOT IsNull(pi) Then
            pi.UninstallProductKey( pi.ProductKeyID)
        End If
    Next 'pi


End Sub 'CleanOSPP

'-------------------------------------------------------------------------------
'   DelSchtasks
'
'   Delete know scheduled tasks.
'-------------------------------------------------------------------------------
Sub DelSchtasks ()
    Dim sCmd

    If CBool(iError AND ERROR_USERCANCEL) Then Exit Sub

    LogH1 "Remove scheduled tasks" 

    LogOnly "FF_INTEGRATEDstreamSchedule"
    oWShell.Run "SCHTASKS /Delete /TN FF_INTEGRATEDstreamSchedule /F", 0, False
    wscript.sleep 500

    LogOnly "FF_INTEGRATEDUPDATEDETECTION"
    oWShell.Run "SCHTASKS /Delete /TN FF_INTEGRATEDUPDATEDETECTION /F", 0, False
    wscript.sleep 500

    LogOnly "C2RAppVLoggingStart"
    oWShell.Run "SCHTASKS /Delete /TN C2RAppVLoggingStart /F", 0, False
    wscript.sleep 500

    LogOnly "Office 15 Subscription Heartbeat"
    sCmd = "SCHTASKS /Delete /TN " & Chr(34) & "Office 15 Subscription Heartbeat" & Chr(34) & " /F"
    oWShell.Run sCmd, 0, False
    wscript.sleep 500

    LogOnly "Microsoft Office 15 Sync Maintenance"
    sCmd = "SCHTASKS /Delete /TN " & Chr(34) & "Microsoft Office 15 Sync Maintenance for {d068b555-9700-40b8-992c-f866287b06c1}" & Chr(34) & " /F"
    oWShell.Run sCmd, 0, False
    wscript.sleep 500

    LogOnly "OfficeInventoryAgentFallBack"
    sCmd = "SCHTASKS /Delete /TN " & Chr(34) & "\Microsoft\Office\OfficeInventoryAgentFallBack" & Chr(34) & " /F"
    oWShell.Run sCmd, 0, False
    wscript.sleep 500

    LogOnly "OfficeTelemetryAgentFallBack"
    sCmd = "SCHTASKS /Delete /TN " & Chr(34) & "\Microsoft\Office\OfficeTelemetryAgentFallBack" & Chr(34) & " /F"
    oWShell.Run sCmd, 0, False
    wscript.sleep 500

    LogOnly "OfficeInventoryAgentLogOn"
    sCmd = "SCHTASKS /Delete /TN " & Chr(34) & "\Microsoft\Office\OfficeInventoryAgentLogOn" & Chr(34) & " /F"
    oWShell.Run sCmd, 0, False

    LogOnly "OfficeTelemetryAgentLogOn"
    sCmd = "SCHTASKS /Delete /TN " & Chr(34) & "\Microsoft\Office\OfficeTelemetryAgentLogOn" & Chr(34) & " /F"
    oWShell.Run sCmd, 0, False

    LogOnly "Office Background Streaming"
    sCmd = "SCHTASKS /Delete /TN " & Chr(34) & "Office Background Streaming" & Chr(34) & " /F"
    oWShell.Run sCmd, 0, False
    wscript.sleep 500

    LogOnly "Office Automatic Updates"
    sCmd = "SCHTASKS /Delete /TN " & Chr(34) & "\Microsoft\Office\Office Automatic Updates" & Chr(34) & " /F"
    oWShell.Run sCmd, 0, False
    wscript.sleep 500

    LogOnly "Office ClickToRun Service Monitor"
    sCmd = "SCHTASKS /Delete /TN " & Chr(34) & "\Microsoft\Office\Office ClickToRun Service Monitor" & Chr(34) & " /F"
    oWShell.Run sCmd, 0, False
    wscript.sleep 500

    LogOnly "Office Subscription Maintenance"
    sCmd = "SCHTASKS /Delete /TN " & Chr(34) & "Office Subscription Maintenance" & Chr(34) & " /F"
    oWShell.Run sCmd, 0, False
    wscript.sleep 500

End Sub

'-------------------------------------------------------------------------------
'   CloseOfficeApps
'
'   End all running instances of applications that will be removed.
'-------------------------------------------------------------------------------
Sub CloseOfficeApps
    Dim Processes, Process, app, prop
    Dim sAppName, sOut, sUserWarn
    Dim fWait
    Dim iRet
    
    On Error Resume Next
    fWait = False
    iProcCloseCnt = iProcCloseCnt + 1
    If fRerun Then Exit Sub

    If NOT fUserConsent Then
        ' detect processes to allow a user warning
        sUserWarn =  "Please save all open documents and close all Office, IE and Windows Explorer applications before proceeding." & vbCrLf & _
                    "When you click OK this removal process will terminate all running Office, IE and Windows Explorer processes and applications." & vbCrLf & vbCrLf & _
                    "Click ‘Cancel’ to to end this removal now."
        For Each app in dicApps.Keys
            sAppName = Replace(app, ".", "%.")
            Set Processes = oWmiLocal.ExecQuery("Select * From Win32_Process Where Name like '" & sAppName & "'")
            For Each Process in Processes
                If NOT InStr(sUserWarn, Process.Name) > 0 Then sUserWarn = sUserWarn & vbCrLf & " - " & Process.Name
            Next 'Process
        Next 'app
        Set Processes = oWmiLocal.ExecQuery("Select * From Win32_Process")
        For Each Process in Processes
            For Each prop in Process.Properties_
                If prop.Name = "ExecutablePath" Then 
                    If IsC2R(prop.Value) Then sUserWarn = sUserWarn & vbCrLf & " - " & Process.Name
                End If 'ExcecutablePath
            Next 'prop
        Next 'Process
        If (InStr(sUserWarn, " - ") > 0 AND NOT fQuiet) Then
            iRet = MsgBox(sUserWarn, 49, "Save your unsaved work now!")
            If iRet = 2 Then 
                SetError ERROR_USERCANCEL
                ExitScript
            Else
                fUserConsent = True
            End If
        End If
    End If 'fUserConsent

    ' end known processes first
    For Each app in dicApps.Keys
        sAppName = Replace(app, ".", "%.")
        Set Processes = oWmiLocal.ExecQuery("Select * From Win32_Process Where Name like '" & sAppName & "'")
        For Each Process in Processes
            sOut = "End process '" & Process.Name
            iRet = Process.Terminate()
            CheckError "CloseOfficeApps: " & Process.Name
            Log sOut & "' returned: " & iRet
            fWait = True
        Next 'Process
    Next 'app

    ' end running applications
    Set Processes = oWmiLocal.ExecQuery("Select * From Win32_Process")
    For Each Process in Processes
        For Each prop in Process.Properties_
            If prop.Name = "ExecutablePath" Then 
                If IsC2R(prop.Value) Then
                    sOut = "End process '" & Process.Name
                    iRet = Process.Terminate()
                    CheckError "CloseOfficeApps: " & Process.Name
                    Log sOut & "' returned: " & iRet
                    fWait = True
                End If 
            End If 'ExcecutablePath
        Next 'prop
    Next 'Process
    If fWait Then wscript.sleep 5000
End Sub 'CloseOfficeApps

'-------------------------------------------------------------------------------
'   Uninstall
'
'   Identify and invoke default uninstall command for a regular uninstall.
'-------------------------------------------------------------------------------
Sub Uninstall
    Dim OseService, srvc
    Dim hDefKey, sSubKeyName, sValue, Name, arrNames, arrTypes
    Dim sku, prod, sUninstallCmd, sReturn, sMsiProp, sCmd
    Dim sPkgFld, sPkgGuid
    Dim i

    If CBool(iError AND ERROR_USERCANCEL) Then Exit Sub

    ' check if OSE service is *installed, *not disabled, *running under System context.
    LogH2 "Check state of OSE service"
    Set OseService = oWmiLocal.Execquery("Select * From Win32_Service Where Name like 'ose%'")
    For Each srvc in OseService
        If (srvc.StartMode = "Disabled") AND (Not srvc.ChangeStartMode("Manual") = 0) Then _
            Log "Conflict detected: OSE service is disabled"
        If (Not srvc.StartName = "LocalSystem") AND (srvc.Change( , , , , , , "LocalSystem", "")) Then _
            Log "Conflict detected: OSE service not running as LocalSystem"
    Next 'srvc

    If NOT dicC2RSuite.Count > 0 Then
        Log "No uninstallable C2R items registered in Uninstall"
    End If

    ' remove the published component registration for C2R packages
    LogH2 "Remove published component registration for C2R packages"
    ' delete the manifest files
    For i = 1 To 4
    	Select Case i
    	Case 1
	    	RegReadValue HKLM, "SOFTWARE\Microsoft\Office\15.0\ClickToRun", "PackageFolder", sPkgFld, "REG_SZ" 
	    	RegReadValue HKLM, "SOFTWARE\Microsoft\Office\15.0\ClickToRun", "PackageGUID", sPkgGuid, "REG_SZ"
	    Case 2
	    	RegReadValue HKLM, "SOFTWARE\Microsoft\Office\16.0\ClickToRun", "PackageFolder", sPkgFld, "REG_SZ" 
	    	RegReadValue HKLM, "SOFTWARE\Microsoft\Office\16.0\ClickToRun", "PackageGUID", sPkgGuid, "REG_SZ"
	    Case 3
	    	RegReadValue HKLM, "SOFTWARE\Microsoft\Office\ClickToRun", "PackageFolder", sPkgFld, "REG_SZ" 
	    	RegReadValue HKLM, "SOFTWARE\Microsoft\Office\ClickToRun", "PackageGUID", sPkgGuid, "REG_SZ"
	    Case 4
	    	sPkgFld = sPackageFolder
	    	sPkgGuid = sPackageGuid
	    End Select
	    If oFso.FolderExists(sValue & "\root\Integration") Then
	        sCmd = "cmd.exe /c del " & chr(34) & sPkgFld & "\root\Integration\C2RManifest*.xml" & chr(34)
	        Log "   Run: " & sCmd
	        sReturn = oWShell.Run (sCmd, 0, True)
	        Log "   Return value: " & sReturn
	        If oFso.FileExists(sPkgFld & "\root\Integration\integrator.exe") Then
	            sCmd = chr(34) & sPkgFld & "\root\Integration\integrator.exe" & chr(34) & " /U  /Extension PackageRoot=" & chr(34) & sPkgFld & "\root" & chr(34) & " PackageGUID=" & sPkgGuid
	            Log "   Run: " & sCmd
	            sReturn = oWShell.Run (sCmd, 0, True)
	            Log "   Return value: " & sReturn
	            sCmd = chr(34) & sPkgFld & "\root\Integration\integrator.exe" & chr(34) & " /U"
	            Log "   Run: " & sCmd
	            sReturn = oWShell.Run (sCmd, 0, True)
	            Log "   Return value: " & sReturn
	        End If
	        If oFso.FileExists(sProgramData & "\Microsoft\ClickToRun\{" & sPkgGuid & "}\integrator.exe") Then
	            sCmd = chr(34) & sProgramData & "\Microsoft\ClickToRun\{" & sPkgGuid & "}\integrator.exe" & chr(34) & " /U  /Extension PackageRoot=" & chr(34) & sPkgFld & "\root" & chr(34) & " PackageGUID=" & sPkgGuid
	            Log "   Run: " & sCmd
	            sReturn = oWShell.Run (sCmd, 0, True)
	            Log "   Return value: " & sReturn
	        End If
	    End If
    Next 'i

    ' delete potential blocking registry keys for msiexec based tasks
    LogH2 "Remove C2R and App-V registry data"
    For Each sku in dicC2RSuite.Keys
        ' remove the ARP entry
        RegDeleteKey HKLM, REG_ARP & sku
    Next 'sku
    RegDeleteKey HKCU, "SOFTWARE\Microsoft\Office\15.0\ClickToRun"
    RegDeleteKey HKLM, "SOFTWARE\Microsoft\Office\15.0\ClickToRun"
    RegDeleteKey HKCU, "SOFTWARE\Microsoft\Office\16.0\ClickToRun"
    RegDeleteKey HKLM, "SOFTWARE\Microsoft\Office\16.0\ClickToRun"
    RegDeleteKey HKCU, "SOFTWARE\Microsoft\Office\ClickToRun"
    RegDeleteKey HKLM, "SOFTWARE\Microsoft\Office\ClickToRun"
    
    ' AppV keys
    hDefKey = HKCU
    sSubKeyName = "SOFTWARE\Microsoft\AppV\ISV"
    Do
        If RegEnumValues(hDefKey, sSubKeyName, arrNames, arrTypes) Then
            For Each Name in arrNames
                If IsC2R(Name) Then RegDeleteValue hDefKey, sSubKeyName, Name, False
            Next 'Name
        End If 'RegEnumValues
        If hDefKey = HKLM Then Exit Do
        hDefKey = HKLM
    Loop

    ' msiexec based uninstall
    sMsiProp = " REBOOT=ReallySuppress NOREMOVESPAWN=True"
    LogH2 "Detect Msi based products"
    For Each prod in oMsi.Products
        If CheckDelete(prod) Then
            Log "Call msiexec.exe to remove " & prod 
            sUninstallCmd = "msiexec.exe /x" & prod & sMsiProp
            If fQuiet Then 
                sUninstallCmd = sUninstallCmd & " /q"
            Else
                sUninstallCmd = sUninstallCmd & " /qb-!"
            End If
            sUninstallCmd = sUninstallCmd & " /l*v " & chr(34) & sLogDir & "\Uninstall_" & prod & ".log" & chr(34)
            CloseOfficeApps
            LogOnly "Call msiexec with '" & sUninstallCmd & "'"
            sReturn = oWShell.Run(sUninstallCmd, 0, True)
            Log "msiexec returned: " & SetupRetVal(sReturn) & " (" & sReturn & ")" & vbCrLf
            fRebootRequired = fRebootRequired OR (sReturn = "3010")
            If fRebootRequired Then SetError ERROR_REBOOT_REQUIRED
            Select Case CInt(sReturn)
            Case ERROR_SUCCESS,ERROR_SUCCESS_CONFIG_COMPLETE,ERROR_SUCCESS_REBOOT_REQUIRED
                'success no action required
            Case Else
                SetError ERROR_STAGE1
            End Select
        Else
        	LogOnly "Skip out of scope product: " & prod
        End If 'CheckDelete
    Next 'Product
    oWShell.Run "cmd.exe /c net stop msiserver", 0, False
End Sub 'Uninstall

'-------------------------------------------------------------------------------
'   RegWipe
'
'   Removal of left behind registry data 
'-------------------------------------------------------------------------------
Sub Regwipe
    Dim hDefKey, item, name, value, RetVal
    Dim sGuid, sSubKeyName, sValue, sCmd
    Dim i, iLoopCnt
    Dim arrKeys, arrNames, arrTypes, arrTestNames, arrTestTypes
    Dim arrMultiSzValues, arrMultiSzNewValues
    Dim fDelReg

    If CBool(iError AND ERROR_USERCANCEL) Then Exit Sub

    LogH1 "Registry CleanUp"
    
    'Moved to earlier timing to avoid reboot needs
    'If NOT dicKeepSku.Count > 0 Then ClearShellIntegrationReg
    
    CloseOfficeApps

    ' Note: ARP entries have already been cleared in uninstall stage

    ' HKCU Registration
    RegDeleteKey HKCU, "Software\Microsoft\Office\15.0\Registration"
    RegDeleteKey HKCU, "Software\Microsoft\Office\16.0\Registration"
    RegDeleteKey HKCU, "Software\Microsoft\Office\Registration"

    
    ' C2R specifics
    ' AppV key "SOFTWARE\Microsoft\AppV" has already been cleared in uninstall stage

    ' Virtual InstallRoot
    RegDeleteKey HKLM, "SOFTWARE\Microsoft\Office\15.0\Common\InstallRoot\Virtual"
    RegDeleteKey HKLM, "SOFTWARE\Microsoft\Office\16.0\Common\InstallRoot\Virtual"
    RegDeleteKey HKLM, "SOFTWARE\Microsoft\Office\Common\InstallRoot\Virtual"

    ' Mapi Search reg
    'O15
    If NOT dicKeepSku.Count > 0 Then RegDeleteKey HKLM, "SOFTWARE\Classes\CLSID\{2027FC3B-CF9D-4ec7-A823-38BA308625CC}"
    'O16
    '{F8E61EDD-EA25-484e-AC8A-7447F2AAE2A9}

    
    ' C2R keys
    RegDeleteKey HKCU, "SOFTWARE\Microsoft\Office\15.0\ClickToRun"
    RegDeleteKey HKLM, "SOFTWARE\Microsoft\Office\15.0\ClickToRun"
    RegDeleteKey HKLM, "SOFTWARE\Microsoft\Office\15.0\ClickToRunStore"
    RegDeleteKey HKCU, "SOFTWARE\Microsoft\Office\16.0\ClickToRun"
    RegDeleteKey HKLM, "SOFTWARE\Microsoft\Office\16.0\ClickToRun"
    RegDeleteKey HKLM, "SOFTWARE\Microsoft\Office\16.0\ClickToRunStore"
    RegDeleteKey HKCU, "SOFTWARE\Microsoft\Office\ClickToRun"
    RegDeleteKey HKLM, "SOFTWARE\Microsoft\Office\ClickToRun"
    RegDeleteKey HKLM, "SOFTWARE\Microsoft\Office\ClickToRunStore"
    
    ' Office key in HKLM
    If Not dicKeepSku.Count > 0 Then
    	'double calls to ensure Wow6432 gets cleared out as well
    	RegDeleteKey HKLM, "Software\Microsoft\Office\15.0"
    	RegDeleteKey HKLM, "Software\Microsoft\Office\15.0"
    	RegDeleteKey HKLM, "Software\Microsoft\Office\16.0"
    	RegDeleteKey HKLM, "Software\Microsoft\Office\16.0"
    End If
    ClearOfficeHKLM "SOFTWARE\Microsoft\Office"

    ' Run key
    sSubKeyName = "SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
    If RegEnumValues (HKLM, sSubKeyName, arrNames, arrTypes) Then
        For Each name in arrNames
            If RegReadValue(HKLM, sSubKeyName, name, sValue, "REG_SZ") Then
                If IsC2R(sValue) Then RegDeleteValue HKLM, sSubKeyName, name, False
            End If
        Next 'item
    End If
    RegDeleteValue HKLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Lync15", False
    RegDeleteValue HKLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Lync16", False
    
    ' ARP
    ' Note: configuration entries have already been removed 
    ' as part of the 'Uninstall' stage
    If RegEnumKey(HKLM, REG_ARP, arrKeys) Then
        For Each item in arrKeys
            If Len(item) > 37 Then
                sGuid = UCase(Left(item, 38))
                If CheckDelete(sGuid) Then RegDeleteKey HKLM, REG_ARP & item & "\"
            End If 'Len(Item)>37
        Next 'Item
    End If

    ' UpgradeCodes, WI config, WI global config
    LogH2 "Scan Windows Installer metadata for removeable UpgradeCodes"
    For iLoopCnt = 1 to 5
        Select Case iLoopCnt
        Case 1
            sSubKeyName = "SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UpgradeCodes\"
            hDefKey = HKLM
        Case 2 
            sSubKeyName = "Installer\UpgradeCodes\"
            hDefKey = HKCR
        Case 3
            sSubKeyName = "SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\"
            hDefKey = HKLM
        Case 4 
            sSubKeyName = "Installer\Features\"
            hDefKey = HKCR
        Case 5 
            sSubKeyName = "Installer\Products\"
            hDefKey = HKCR
        End Select
        If RegEnumKey(hDefKey, sSubKeyName, arrKeys) Then
            For Each item in arrKeys
                ' ensure the expected length for a compressed GUID
                If Len(item) = 32 Then
                    ' expand the GUID
                    sGuid = GetExpandedGuid(item) 
                    ' check if it's an Office key
                    If CheckDelete(sGuid) Then
                        If iLoopCnt < 3 Then
                            ' enum all entries
                            RegEnumValues hDefKey, sSubKeyName & item, arrNames, arrTypes
                            If IsArray(arrNames) Then
                                ' delete entries within removal scope
                                For Each name in arrNames
                                    If Len(name) = 32 Then
                                        sGuid = GetExpandedGuid(name)
                                        If CheckDelete(sGuid) Then RegDeleteValue hDefKey, sSubKeyName & item & "\", name, True
                                    Else
                                        ' invalid data -> delete the value
                                        RegDeleteValue hDefKey, sSubKeyName & item & "\", name, True
                                    End If
                                Next 'Name
                            End If 'IsArray(arrNames)
                            ' if all entries were removed - delete the key
                            If NOT RegEnumValues(hDefKey, sSubKeyName & item, arrNames, arrTypes) Then RegDeleteKey hDefKey, sSubKeyName & item & "\"
                        Else 'iLoopCnt >= 3
                            RegDeleteKey hDefKey, sSubKeyName & item & "\"
                        End If 'iLoopCnt < 3
                    End If 'InScope
                End If 'Len(Item)=32
            Next 'Item
        End If 'RegEnumKey
    Next 'iLoopCnt

    ' Components in Global
    LogH2 "Scan Windows Installer Global Components metadata"
    sSubKeyName = "SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\"
    hDefKey = HKLM
    If RegEnumKey(hDefKey, sSubKeyName, arrKeys) Then
        For Each item in arrKeys
            ' ensure the expected length for a compressed GUID
            If Len(Item) = 32 Then
                If RegEnumValues(hDefKey, sSubKeyName & item, arrNames, arrTypes) Then
                    For Each name in arrNames
                        If Len(Name) = 32 Then
                            sGuid = GetExpandedGuid(Name)
                            If CheckDelete(sGuid) Then
                                RegDeleteValue hDefKey, sSubKeyName & item & "\", name, False
                                ' if all entries were removed - delete the key
                                If NOT RegEnumValues(hDefKey, sSubKeyName & item, arrTestNames, arrTestTypes) Then RegDeleteKey hDefKey, sSubKeyName & item & "\"
                            End If
                        End If '32
                    Next 'Name
                End If 'RegEnumValues
            End If '32
        Next 'Item
    End If 'RegEnumKey

    ' Published Components
    LogH2 "Scanning Windows Installer Published Components metadata"
    sSubKeyName = "Installer\Components\"
    hDefKey = HKCR
    If RegEnumKey(hDefKey, sSubKeyName, arrKeys) Then
        For Each item in arrKeys
            ' ensure the expected length for a compressed GUID
            If Len(Item) = 32 Then
                If RegEnumValues(hDefKey, sSubKeyName & item, arrNames, arrTypes) Then
                    For Each name in arrNames
                        If RegReadValue (hDefKey, sSubKeyName & item, name, sValue, "REG_MULTI_SZ") Then
                            arrMultiSzValues = Split(sValue, chr(13))
                            If IsArray(arrMultiSzValues) Then
                                i = -1
                                ReDim arrMultiSzNewValues(-1)
                                fDelReg = False
                                For Each value in arrMultiSzValues
                                    If Len(value) > 19 Then
                                        sGuid = ""
                                        If GetDecodedGuid(Left(value, SQUISHED), sGuid) Then
                                            If CheckDelete(sGuid) Then
                                                fDelReg = True
                                            Else
                                                i = i + 1 
                                                ReDim Preserve arrMultiSzNewValues(i)
                                                arrMultiSzNewValues(i) = value
                                            End If 'CheckDelete
                                        End If 'decode
                                    End If '19
                                Next 'Value
                                If NOT (i = -1) Then
                                    If NOT UBound(arrMultiSzValues) = i Then oReg.SetMultiStringValue hDefKey, sSubKeyName & item, name,arrMultiSzNewValues
                                Else
                                    If fDelReg Then
                                        RegDeleteValue hDefKey, sSubKeyName & item & "\", name, True
                                        ' if all entries were removed - delete the key
                                        If NOT RegEnumValues(hDefKey, sSubKeyName & item, arrTestNames, arrTestTypes) Then RegDeleteKey hDefKey, sSubKeyName & item & "\"
                                    End If 'DelReg
                                End If
                            End If 'IsArray
                        End If
                    Next 'Name
                End If 'RegEnumValues
            End If '32
        Next 'Item
    End If 'RegEnumKey

End Sub 'Regwipe

'-------------------------------------------------------------------------------
'   ClearShellIntegrationReg
'
'   Delete registry items that may cause Explorer / Windows Shell to have a lock
'   on files 
'-------------------------------------------------------------------------------
Sub ClearShellIntegrationReg
    Dim Processes, Process
    Dim sOut
    Dim iRet
    
    Set Processes = oWmiLocal.ExecQuery("Select * From Win32_Process Where Name like 'explorer.exe'")
    For Each Process in Processes
        sOut = "End process '" & Process.Name
        iRet = Process.Terminate()
        CheckError "CloseOfficeApps: " & Process.Name
        Log sOut & "' returned: " & iRet
    Next 'Process
    wscript.sleep 500

    
    ' Protocol Handlers
    RegDeleteKey HKLM, "SOFTWARE\Classes\Protocols\Handler\osf"

    ' Groove ShellIconOverlayIdentifiers
    RegDeleteKey HKLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Microsoft SPFS Icon Overlay 1 (ErrorConflict)"
    RegDeleteKey HKLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Microsoft SPFS Icon Overlay 2 (SyncInProgress)"
    RegDeleteKey HKLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Microsoft SPFS Icon Overlay 3 (InSync)"
    RegDeleteKey HKLM, "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Microsoft SPFS Icon Overlay 1 (ErrorConflict)"
    RegDeleteKey HKLM, "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Microsoft SPFS Icon Overlay 2 (SyncInProgress)"
    RegDeleteKey HKLM, "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Microsoft SPFS Icon Overlay 3 (InSync)"

    ' Shell extensions
    RegDeleteValue HKLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\", "{B28AA736-876B-46DA-B3A8-84C5E30BA492}", False
    RegDeleteValue HKLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\", "{8B02D659-EBBB-43D7-9BBA-52CF22C5B025}", False
    RegDeleteValue HKLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\", "{0875DCB6-C686-4243-9432-ADCCF0B9F2D7}", False
    RegDeleteValue HKLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\", "{42042206-2D85-11D3-8CFF-005004838597}", False
    RegDeleteValue HKLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\", "{993BE281-6695-4BA5-8A2A-7AACBFAAB69E}", False
    RegDeleteValue HKLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\", "{C41662BB-1FA0-4CE0-8DC5-9B7F8279FF97}", False
    RegDeleteValue HKLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\", "{506F4668-F13E-4AA1-BB04-B43203AB3CC0}", False
    RegDeleteValue HKLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\", "{D66DC78C-4F61-447F-942B-3FB6980118CF}", False
    RegDeleteValue HKLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\", "{46137B78-0EC3-426D-8B89-FF7C3A458B5E}", False
    RegDeleteValue HKLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\", "{8BA85C75-763B-4103-94EB-9470F12FE0F7}", False
    RegDeleteValue HKLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\", "{CD55129A-B1A1-438E-A425-CEBC7DC684EE}", False
    RegDeleteValue HKLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\", "{D0498E0A-45B7-42AE-A9AA-ABA463DBD3BF}", False
    RegDeleteValue HKLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\", "{E768CD3B-BDDC-436D-9C13-E1B39CA257B1}", False

    ' BHO
    RegDeleteKey HKLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{31D09BA0-12F5-4CCE-BE8A-2923E76605DA}"
    RegDeleteKey HKLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{B4F3A835-0E21-4959-BA22-42B3008E02FF}"
    RegDeleteKey HKLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{D0498E0A-45B7-42AE-A9AA-ABA463DBD3BF}"
    RegDeleteKey HKLM, "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{31D09BA0-12F5-4CCE-BE8A-2923E76605DA}"
    RegDeleteKey HKLM, "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{B4F3A835-0E21-4959-BA22-42B3008E02FF}"
    RegDeleteKey HKLM, "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{D0498E0A-45B7-42AE-A9AA-ABA463DBD3BF}"

    ' OneNote Namespace Extension for Desktop
    RegDeleteKey HKLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\{0875DCB6-C686-4243-9432-ADCCF0B9F2D7}"

    ' Web Sites
    RegDeleteKey HKLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\Namespace\{B28AA736-876B-46DA-B3A8-84C5E30BA492}"
    RegDeleteKey HKLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\NetworkNeighborhood\Namespace\{46137B78-0EC3-426D-8B89-FF7C3A458B5E}"

    ' VolumeCaches
    RegDeleteKey HKLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Microsoft Office Temp Files"

    Set Processes = oWmiLocal.ExecQuery("Select * From Win32_Process Where Name like 'explorer.exe'")
    For Each Process in Processes
        sOut = "End process '" & Process.Name
        iRet = Process.Terminate()
        CheckError "CloseOfficeApps: " & Process.Name
        Log sOut & "' returned: " & iRet
    Next 'Process
    wscript.sleep 500
    RestoreExplorer

End Sub 'ClearShellIntegrationReg

'-------------------------------------------------------------------------------
'   FileWipe
'
'   Removal of left behind services, files and shortcuts 
'-------------------------------------------------------------------------------
Sub FileWipe
    Dim scRoot
    Dim fDelFolders
    
    If CBool(iError AND ERROR_USERCANCEL) Then Exit Sub

    LogH1 "File Cleanup" 

    fDelFolders = False
    CloseOfficeApps
    DelSchtasks

    LogH1 "Delete Services"
    ' remove the OfficeSvc service
    LogH2 "Delete OfficeSvc service"
    DeleteService "OfficeSvc"

    ' SP1 addition / change
    ' remove the ClickToRunSvc service
    LogH2 "Delete ClickToRunSvc service" 
    DeleteService "ClickToRunSvc"

    ' adding additional processes for termination
    'dicApps.Add "explorer.exe", "explorer.exe"
    dicApps.Add "msiexec.exe", "msiexec.exe"
    dicApps.Add "ose.exe", "ose.exe"
    
    If fC2R Then
	    LogH1 "Delete Files and Folders"
        ' delete C2R package files
        LogH2 "Delete C2R package files" 
        If oFso.FolderExists(sProgramFiles & "\Microsoft Office 15") _
        Or oFso.FolderExists(sProgramFiles & "\Microsoft Office 16") _
        Or oFso.FolderExists(oWShell.ExpandEnvironmentStrings("%programfiles%") & "\Microsoft Office\PackageManifests") _
        Or oFso.FolderExists(oWShell.ExpandEnvironmentStrings("%programfiles(x86)%") & "\Microsoft Office\PackageManifests") Then
            fDelFolders = True
            'Log "   Attention: Now closing Explorer.exe for file delete operations"
            'Log "   Explorer will automatically restart."
            wscript.sleep 2000
            CloseOfficeApps
        End If
        ' delete Office folders
        LogH2 "Delete Office folders"
        DeleteFolder sProgramFiles & "\Microsoft Office 15"
        DeleteFolder sProgramFiles & "\Microsoft Office 16"
        If f64 Then 
        	DeleteFolder sCommonProgramFilesX86 & "\Microsoft Office 15"
        	DeleteFolder sCommonProgramFilesX86 & "\Microsoft Office 16"
        End If
        If fDelFolders Then
        	DeleteFolder sProgramFiles & "\Microsoft Office\PackageManifests"
        	DeleteFolder sProgramFiles & "\Microsoft Office\PackageSunrisePolicies"
        	DeleteFolder sProgramFiles & "\Microsoft Office\root"
        	DeleteFile sProgramFiles & "\Microsoft Office\AppXManifest.xml"
         	DeleteFile sProgramFiles & "\Microsoft Office\FileSystemMetadata.xml"
        	If Not dicKeepSku.Count > 0 Then 
        		DeleteFolder sProgramFiles & "\Microsoft Office\Office16"
        		DeleteFolder sProgramFiles & "\Microsoft Office\Office15"
        	End If
        	If f64 Then
	        	DeleteFolder sProgramFilesX86 & "\Microsoft Office\PackageManifests"
	        	DeleteFolder sProgramFilesX86 & "\Microsoft Office\PackageSunrisePolicies"
	        	DeleteFolder sProgramFilesX86 & "\Microsoft Office\root"
	        	DeleteFile sProgramFilesX86 & "\Microsoft Office\AppXManifest.xml"
	         	DeleteFile sProgramFilesX86 & "\Microsoft Office\FileSystemMetadata.xml"
	        	If Not dicKeepSku.Count > 0 Then 
	        		DeleteFolder sProgramFilesX86 & "\Microsoft Office\Office16"
	        		DeleteFolder sProgramFilesX86 & "\Microsoft Office\Office15"
	        	End If
       		End If
		End If
        
        DeleteFolder sProgramData & "\Microsoft\ClickToRun"
        DeleteFolder sCommonProgramFiles & "\microsoft shared\ClickToRun"
        DeleteFolder sProgramData & "\Microsoft\office\FFPackageLocker"
        DeleteFolder sProgramData & "\Microsoft\office\ClickToRunPackageLocker"
        If oFso.FileExists(sProgramData & "\Microsoft\office\FFPackageLocker") Then DeleteFile sProgramData & "\Microsoft\office\FFPackageLocker"
        If oFso.FileExists(sProgramData & "\Microsoft\office\FFStatePBLocker") Then DeleteFile sProgramData & "\Microsoft\office\FFStatePBLocker"
        If NOT dicKeepSku.Count > 0 Then DeleteFolder sProgramData & "\Microsoft\office\Heartbeat"
        DeleteFolder oWShell.ExpandEnvironmentStrings("%userprofile%") & "\Microsoft Office"
        DeleteFolder oWShell.ExpandEnvironmentStrings("%userprofile%") & "\Microsoft Office 15"
        DeleteFolder oWShell.ExpandEnvironmentStrings("%userprofile%") & "\Microsoft Office 16"
    End If

    ' restore explorer.exe if needed
    RestoreExplorer

    ' delete shortcuts
    LogH2 "Search and delete shortcuts"
    CleanShortcuts sAllUsersProfile, True, False
    CleanShortcuts sProfilesDirectory, True, False

    ' delete empty folder structures
    If dicDelFolder.Count > 0 Then
        LogH2 "Remove empty folders"
        DeleteEmptyFolders
    End If

    ' add the collected files in use for delete on reboot
    If dicDelInUse.Count > 0 Then ScheduleDeleteEx

    LogH2 "File Cleanup complete"
End Sub ' FileWipe

'-------------------------------------------------------------------------------
'   CleanShortcuts
'
'   Recursively search all profile folders for Office shortcuts in scope 
'-------------------------------------------------------------------------------
Sub CleanShortcuts (sFolder, fDelete, fUnPin)
    Dim oFolder, fld, file, sc, item
    Dim fDeleteSC

	If fSkipSD Then Exit Sub
	
	Set oFolder = oFso.GetFolder(sFolder)
	' exclude system protected link folders
    If CBool(oFolder.Attributes AND 1024) Then Exit Sub

    On Error Resume Next
    For Each fld In oFolder.SubFolders
        If Err <> 0 Then
		    CheckError "CleanShortcuts: " & vbTab & sFolder
        Else
            CleanShortcuts fld.Path, fDelete, fUnPin
        End If
	Next
    For Each file In oFolder.Files
		If LCase(Right(file.Path, 4)) = ".lnk" Then
            fDeleteSC = False
            LogOnly " check file: " & file.Path
            set sc = oWShell.CreateShortcut(file.Path)
            If Err <> 0 Then
		        CheckError "CleanShortcutsSC: " & vbTab & sFolder
            Else
                'Compare if the shortcut target is in the list of executables that will be removed
                'LogOnly "  - SC.TargetPath: " & sc.TargetPath
                If Len(sc.TargetPath) > 0 Then
                    If InStr(sc.TargetPath,"{") > 0 Then
                        'Handle Windows Installer shortcuts
                        If Len(sc.TargetPath) >= InStr(sc.TargetPath,"{") + 37 Then
                            If CheckDelete(Mid(sc.TargetPath, InStr(sc.TargetPath,"{"), 38)) Then fDeleteSC = True
                        End If
                    Else
                        'Handle regular shortcuts
                        If IsC2R(sc.TargetPath) Then fDeleteSC = True
                        If NOT oFso.FileExists(sc.TargetPath) Then
                            ' Shortcut target does not exist
                            If IsC2R(sc.TargetPath) Then
                                LogOnly "remove Office shortcut with non-existent target: " & file.Path & " - " & sc.TargetPath
                                fDeleteSC = True
                            Else
                                'LogOnly "  - keep orphaned SC as target is not in scope: " & sc.TargetPath
                            End If
                        Else
                            'LogOnly "  - keep SC as shortcut target does still exist: " & sc.TargetPath
                        End If
                    End If
                End If
            End If
            If fDeleteSC Then 
                If NOT dicDelFolder.Exists(sFolder) Then dicDelFolder.Add sFolder, sFolder
                If fUnPin OR fDelete Then 
                    If oFso.FileExists(sc.TargetPath) Then
                        UnPin file
                    Else
                        sc.TargetPath = sNotepad
                        sc.Save
                        UnPin file
                    End If
                End If
                If fDelete Then DeleteFile file.Path
                fDeleteSC = False
                fClearTaskBand = True
            End If 'fDeleteSC
        End If
	Next
    On Error Goto 0
End Sub 'CleanShortcuts

'-------------------------------------------------------------------------------
'   UnPin
'
'   Unpins a shortcut from the taskbar or start menu 
'-------------------------------------------------------------------------------
Sub UnPin(file)
    Dim fldItem, verb

    On Error Resume Next
    Set fldItem = oShellApp.NameSpace(file.ParentFolder.Path).ParseName(file.Name)
    For Each verb in fldItem.Verbs
        Select Case LCase(Replace(verb, "&", ""))
        Case "unpin from taskbar", "von taskleiste lösen", "détacher du barre des tâches", "détacher de la barre des tâches", "desanclar de la barra de tareas", "ta bort från aktivitetsfältet", "frigør fra proceslinje", "frigør fra proceslinjen", "desanclar de la barra de tareas", "odepnout z hlavního panelu", "van de taakbalk losmaken", "poista kiinnitys tehtäväpalkista", "rimuovi dalla barra delle applicazioni"
            LogOnly "unpin Office shortcut from taskbar: " & file.Name 
            verb.DoIt
        Case "unpin from start menu", "vom startmenü lösen", "désépingler du menu démarrer", "supprimer du menu démarrer", "détacher du menu démarrer", "détacher de la menu démarrer", "odepnout z nabídky start", "frigør fra menuen start", "van het menu start losmaken", "losmaken van menu start", "poista kiinnitys käynnistä-valikosta", "irrota aloitusvalikosta"
            LogOnly "unpin Office shortcut from start menu: " & file.Name 
            If iVersionNT > 600 Then verb.DoIt
        End Select
        Select Case Replace(verb, "&", "")
        Case "从「开始」菜单解锁", "從 [開始] 功能表取消釘選", "タスク バーに表示しない(K)", "작업 표시줄에서 제거(K)", "Открепить от панели задач", "Ξεκαρφίτσωμα από το μενού Έναρξη", "‏‏בטל הצמדה לתפריט התחלה"
            LogOnly "unpin Office shortcut: " & file.Name 
            verb.DoIt
        End Select
    Next
    On Error Goto 0
End Sub

'-------------------------------------------------------------------------------
'   ClearTaskBand
'
'   Clears contents from the users taskband to get rid of pinned items
'-------------------------------------------------------------------------------
Sub ClearTaskBand ()
    Dim sid
    Dim sTaskBand, sHKUTaskBand
    Dim arrSid

    sTaskBand = "Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband\"
    RegDeleteValue HKCU, sTaskBand, "Favorites", False
    RegDeleteValue HKCU, sTaskBand, "FavoritesRemovedChanges", False
    RegDeleteValue HKCU, sTaskBand, "FavoritesChanges", False
    RegDeleteValue HKCU, sTaskBand, "FavoritesResolve", False
    RegDeleteValue HKCU, sTaskBand, "FavoritesVersion", False

    ' enum all profiles in HKU
    LoadUsersReg
    If NOT RegEnumKey(HKU, "", arrSid) Then Exit Sub
    For Each sid in arrSid
        sHKUTaskBand = sid & "\" & sTaskBand
        RegDeleteValue HKCU, sHKUTaskBand, "Favorites", False
        RegDeleteValue HKCU, sHKUTaskBand, "FavoritesRemovedChanges", False
        RegDeleteValue HKCU, sHKUTaskBand, "FavoritesChanges", False
        RegDeleteValue HKCU, sHKUTaskBand, "FavoritesResolve", False
        RegDeleteValue HKCU, sHKUTaskBand, "FavoritesVersion", False
    Next 'sid
End Sub 'ClearTaskBand

'-------------------------------------------------------------------------------
'   LoadUsersReg
'
'   Loads the HKCU for all local users
'-------------------------------------------------------------------------------
Sub LoadUsersReg ()
    Dim profilefolder
    Dim sValue

    LogH1 "Load User Registry Profiles"
    On Error Resume Next

    oReg.GetExpandedStringValue HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList", "ProfilesDirectory", sValue
    For Each profilefolder in oFso.GetFolder(sValue).SubFolders
        If oFso.FileExists(profilefolder.path & "\ntuser.dat") Then
            LogOnly " load: " & profilefolder.path & "\ntuser.dat" & " as " & "HKU\" & profilefolder.name
            oWShell.Run "reg load " & _
                                    chr(34) & "HKU\" & profilefolder.name & chr(34) & " " & _
                                    chr(34) & profilefolder.path & "\ntuser.dat" & chr(34), 0, True
        End If
'        If oFso.FileExists(profilefolder.path & "\Local Settings\Application Data\Microsoft\Windows\UsrClass.dat") Then
'            LogOnly " load: " & profilefolder.path & "\..\UsrClass.dat" & " as " & "HKU\" & profilefolder.name & "_Classes"
'            oWShell.Run "reg load " & _
'                                    chr(34) & "HKU\" & profilefolder.name & "_Classes" & chr(34) & " " & _
'                                    chr(34) & profilefolder.path & "\Local Settings\Application Data\Microsoft\Windows\UsrClass.dat" & chr(34),0,True
'        End If
    Next
End Sub

'-------------------------------------------------------------------------------
'   ClearOfficeHKLM
'
'   Recursively search and clear the HKLM Office key from references in scope 
'-------------------------------------------------------------------------------
Sub ClearOfficeHKLM (sSubKeyName)
    Dim key, name
    Dim sValue
    Dim arrKeys, arrNames, arrTypes
    Dim arrTestNames, arrTestTypes, arrTestKeys

    ' recursion
    If RegEnumKey(HKLM, sSubKeyName, arrKeys) Then
        For Each key in arrKeys
            ClearOfficeHKLM sSubKeyName & "\" & key
        Next 'key
    End If
    
    ' identify & clear removable entries
    If RegEnumValues(HKLM, sSubKeyName, arrNames, arrTypes) Then
        For Each name in arrNames
            If RegReadValue(HKLM, sSubKeyName, name, sValue, "REG_SZ") Then
                If IsC2R(sValue) Then RegDeleteValue HKLM, sSubKeyName, name, False
            End If
        Next 'item
    End If
    
    ' clear out empty keys
    If (NOT RegEnumValues(HKLM, sSubKeyName, arrNames, arrTypes)) AND _
       (NOT RegEnumKey(HKLM, sSubKeyName, arrKeys)) AND _
       (NOT dicKeepSku.Count > 0) Then _
        RegDeleteKey HKLM, sSubKeyName
End Sub


'-------------------------------------------------------------------------------
'
'                                        Helper Functions
'
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'   IsC2R
'
'   Check if the passed in string is related to C2R
'   Returns TRUE if in C2R scope
'-------------------------------------------------------------------------------
Function IsC2R (sValue)

	Const OREF            = "\ROOT\OFFICE1"
	Const OREFROOT        = "Microsoft Office\Root\"
	Const OREGREFC2R15    = "Microsoft Office 15"
	Const OREGREFC2R16    = "Microsoft Office 16"
	Const OCOMMON		  = "\microsoft shared\ClickToRun"
	Const OMANIFEST		  = "\Microsoft Office\PackageManifests"
	Const OSUNRISE		  = "\Microsoft Office\PackageSunrisePolicies"
	
	Dim fReturn
	
	fReturn = False
	
    If InStr(LCase(sValue), LCase(OREF)) > 0 _
    Or InStr(LCase(sValue), LCase(OREFROOT)) > 0 _
    Or InStr(LCase(sValue), LCase(OCOMMON)) > 0 _
    Or InStr(LCase(sValue), LCase(OMANIFEST)) > 0 _
    Or InStr(LCase(sValue), LCase(OSUNRISE)) > 0 _
    Or InStr(LCase(sValue), LCase(OREGREFC2R15)) > 0 _
    Or InStr(LCase(sValue), LCase(OREGREFC2R16)) > 0 Then fReturn = True
    
	IsC2R = fReturn
End Function

'-------------------------------------------------------------------------------
'   CheckRegPermissions
'
'   Test the permissions on some key registry locations to determine if 
'   sufficient permissions are given.
'-------------------------------------------------------------------------------
Function CheckRegPermissions
    Const KEY_QUERY_VALUE       = &H0001
    Const KEY_SET_VALUE         = &H0002
    Const KEY_CREATE_SUB_KEY    = &H0004
    Const DELETE                = &H00010000

    Dim sSubKeyName
    Dim fReturn

    CheckRegPermissions = True
    sSubKeyName = "Software\Microsoft\Windows\"
    oReg.CheckAccess HKLM, sSubKeyName, KEY_QUERY_VALUE, fReturn
    If Not fReturn Then CheckRegPermissions = False
    oReg.CheckAccess HKLM, sSubKeyName, KEY_SET_VALUE, fReturn
    If Not fReturn Then CheckRegPermissions = False
    oReg.CheckAccess HKLM, sSubKeyName, KEY_CREATE_SUB_KEY, fReturn
    If Not fReturn Then CheckRegPermissions = False
    oReg.CheckAccess HKLM, sSubKeyName, DELETE, fReturn
    If Not fReturn Then CheckRegPermissions = False

End Function 'CheckRegPermissions

'-------------------------------------------------------------------------------
'   GetMyProcessId
'
'   Returns the process id of the own process
'-------------------------------------------------------------------------------
Function GetMyProcessId()
    Dim iParentProcessId

    iParentProcessId = 0
    ' try to obtain from creating a new cscript instance
    On Error Resume Next
    iParentProcessId = GetObject("winmgmts:root\cimv2").Get("Win32_Process.Handle='" & oWShell.Exec("cscript.exe").ProcessId & "'").ParentProcessId
    On Error Goto 0
    If iParentProcessId > 0 Then
        ' succeeded to obtain the process id
        GetMyProcessId = iParentProcessId
        Exit Function
    End If

    ' failed to obtain the id from the creation of a new instance
    ' get it from enum of Win32_Process
    Dim Process, Processes
    Err.Clear
    Set Processes = oWmiLocal.ExecQuery("Select * From Win32_Process WHERE Name='cscript.exe' AND CommandLine like '%" & SCRIPTNAME & "%'")
    For Each Process in Processes
        iParentProcessId = Process.ProcessId
        Exit For
    Next
    GetMyProcessId = iParentProcessId
End Function 'GetMyProcessId

'-------------------------------------------------------------------------------
'   Delimiter
'
'   Returns the delimiter for a passed in string
'-------------------------------------------------------------------------------
Function Delimiter (sVersion)
    Dim iCnt, iAsc

    Delimiter = " "
    For iCnt = 1 To Len(sVersion)
        iAsc = Asc(Mid(sVersion, iCnt, 1))
        If Not (iASC >= 48 And iASC <= 57) Then 
            Delimiter = Mid(sVersion, iCnt, 1)
            Exit Function
        End If
    Next 'iCnt
End Function

'-------------------------------------------------------------------------------
'   GetExpandedGuid
'
'   Returns the expanded string from a compressed GUID
'-------------------------------------------------------------------------------
Function GetExpandedGuid (sGuid)
    Dim i

    'Ensure valid length
    If NOT Len(sGuid) = 32 Then Exit Function

    GetExpandedGuid = "{" & StrReverse(Mid(sGuid,1,8)) & "-" & _
                       StrReverse(Mid(sGuid,9,4)) & "-" & _
                       StrReverse(Mid(sGuid,13,4))& "-"
    For i = 17 To 20
	    If i Mod 2 Then
		    GetExpandedGuid = GetExpandedGuid & mid(sGuid,(i + 1),1)
	    Else
		    GetExpandedGuid = GetExpandedGuid & mid(sGuid,(i - 1),1)
	    End If
    Next
    GetExpandedGuid = GetExpandedGuid & "-"
    For i = 21 To 32
	    If i Mod 2 Then
		    GetExpandedGuid = GetExpandedGuid & mid(sGuid,(i + 1),1)
	    Else
		    GetExpandedGuid = GetExpandedGuid & mid(sGuid,(i - 1),1)
	    End If
    Next
    GetExpandedGuid = GetExpandedGuid & "}"
End Function 'GetExpandedGuid

'-------------------------------------------------------------------------------
'   GetCompressedGuid
'
'   Returns the compressed string for a GUID
'-------------------------------------------------------------------------------
Function GetCompressedGuid (sGuid)
    Dim sCompGUID
    Dim i
    
    'Ensure Valid Length
    If NOT Len(sGuid) = 38 Then Exit Function

    sCompGUID = StrReverse(Mid(sGuid,2,8))  & _
                StrReverse(Mid(sGuid,11,4)) & _
                StrReverse(Mid(sGuid,16,4)) 
    For i = 21 To 24
	    If i Mod 2 Then
		    sCompGUID = sCompGUID & Mid(sGuid, (i + 1), 1)
	    Else
		    sCompGUID = sCompGUID & Mid(sGuid, (i - 1), 1)
	    End If
    Next
    For i = 26 To 37
	    If i Mod 2 Then
		    sCompGUID = sCompGUID & Mid(sGuid, (i - 1), 1)
	    Else
		    sCompGUID = sCompGUID & Mid(sGuid, (i + 1), 1)
	    End If
    Next
    GetCompressedGuid = sCompGUID
End Function

'-------------------------------------------------------------------------------
'   GetDecodedGuid
'
'   Returns the GUID from a squished format
'-------------------------------------------------------------------------------
Function GetDecodedGuid(sEncGuid, sGuid)

Dim sDecode, sTable, sHex, iChr
Dim arrTable
Dim i, iAsc, pow85, decChar
Dim lTotal
Dim fFailed

    fFailed = False

    sTable =    "0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff," & _
                "0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff," & _
                "0xff,0x00,0xff,0xff,0x01,0x02,0x03,0x04,0x05,0x06,0x07,0x08,0x09,0x0a,0x0b,0xff," & _
                "0x0c,0x0d,0x0e,0x0f,0x10,0x11,0x12,0x13,0x14,0x15,0xff,0xff,0xff,0x16,0xff,0x17," & _
                "0x18,0x19,0x1a,0x1b,0x1c,0x1d,0x1e,0x1f,0x20,0x21,0x22,0x23,0x24,0x25,0x26,0x27," & _
                "0x28,0x29,0x2a,0x2b,0x2c,0x2d,0x2e,0x2f,0x30,0x31,0x32,0x33,0xff,0x34,0x35,0x36," & _
                "0x37,0x38,0x39,0x3a,0x3b,0x3c,0x3d,0x3e,0x3f,0x40,0x41,0x42,0x43,0x44,0x45,0x46," & _
                "0x47,0x48,0x49,0x4a,0x4b,0x4c,0x4d,0x4e,0x4f,0x50,0x51,0x52,0xff,0x53,0x54,0xff"
    arrTable = Split(sTable,",")
    lTotal = 0 : pow85 = 1
    For i = 0 To 19
        fFailed = True
        If i Mod 5 = 0 Then
            lTotal = 0 : pow85 = 1
        End If ' i Mod 5 = 0
        iAsc = Asc(Mid(sEncGuid,i+1,1))
        sHex = arrTable(iAsc)
        If iAsc >=128 Then Exit For
        If sHex = "0xff" Then Exit For
        iChr = CInt("&h"&Right(sHex,2))
        lTotal = lTotal + (iChr * pow85)
        If i Mod 5 = 4 Then sDecode = sDecode & DecToHex(lTotal)
        pow85 = pow85 * 85
        fFailed = False
    Next 'i
    If NOT fFailed Then sGuid = "{"&Mid(sDecode,1,8)&"-"& _
                                Mid(sDecode,13,4)&"-"& _
                                Mid(sDecode,9,4)&"-"& _
                                Mid(sDecode,23,2) & Mid(sDecode,21,2)&"-"& _
                                Mid(sDecode,19,2) & Mid(sDecode,17,2) & Mid(sDecode,31,2) & Mid(sDecode,29,2) & Mid(sDecode,27,2) & Mid(sDecode,25,2) &"}"

    GetDecodedGuid = NOT fFailed

End Function 'GetDecodedGuid

'-------------------------------------------------------------------------------
'   DecToHex
'
'   Convert a long decimal to hex
'-------------------------------------------------------------------------------
Function DecToHex(lDec)
    
    Dim sHex
    Dim iLen
    Dim lVal, lExp
    Dim arrChr
  
    arrChr = Array("0","1","2","3","4","5","6","7","8","9","A","B","C","D","E","F")
    sHex = ""
    lVal = lDec
    lExp = 16^10
    While lExp >= 1
        If lVal >= lExp Then
            sHex = sHex & arrChr(Int(lVal / lExp))
            lVal = lVal - lExp * Int(lVal / lExp)
        Else
            sHex = sHex & "0"
            If sHex = "0" Then sHex = ""
        End If
        lExp = lExp / 16
    Wend

    iLen = 8 - Len(sHex)
    If iLen > 0 Then sHex = String(iLen, "0") & sHex
    DecToHex = sHex
End Function

'-------------------------------------------------------------------------------
'   RelaunchAs64Host
'
'   Relaunch self with 64 bit CScript host
'-------------------------------------------------------------------------------
Sub RelaunchAs64Host
    Dim Argument, sCmd
    Dim fQuietRelaunch

    fQuietRelaunch = False
    sCmd = Replace(LCase(wscript.Path), "syswow64", "sysnative") & "\cscript.exe " & Chr(34) & WScript.scriptFullName & Chr(34)
    If fQuiet Then fQuietRelaunch = True
    If Wscript.Arguments.Count > 0 Then
        For Each Argument in Wscript.Arguments
            sCmd = sCmd  &  " " & chr(34) & Argument & chr(34)
            Select Case UCase(Argument)
            Case "/Q", "/QUIET"
                fQuietRelaunch = True
            End Select
        Next 'Argument
    End If
    sCmd = sCmd & " /ChangedHostBitness"
    If fQuietRelaunch Then
        sCmd = Replace (sCmd, "\cscript.exe", "\wscript.exe")
        Wscript.Quit CLng(oWShell.Run (sCmd, 0, True))
    Else
        Wscript.Quit CLng(oWShell.Run (sCmd, 1, True))
    End If

End Sub 'RelaunchAs64Host

'-------------------------------------------------------------------------------
'   RelaunchElevated
'
'   Relaunch the script with elevated permissions
'-------------------------------------------------------------------------------
Sub RelaunchElevated
    Dim Argument, Process, Processes
    Dim iParentProcessId, iSpawnedProcessId
    Dim sCmdLine, sRetValFile, sValue
    Dim oShell

    SetError ERROR_RELAUNCH
    ' Shell object for relaunch
    Set oShell = CreateObject("Shell.Application")
    ' Note: Command line has not been parsed at this point
    ' build command line for relaunch
    sCmdLine = Chr(34) & WScript.ScriptFullName & Chr(34)
    If Wscript.Arguments.Count > 0 Then
        For Each Argument in Wscript.Arguments
            Select Case UCase(Argument)
            Case "/Q","/QUIET"
                'Don't try to relaunch in quiet mode
                Exit Sub
                SetError ERROR_ELEVATION_FAILED
            Case "UAC"
                'Already tried elevated relaunch
                SetError ERROR_ELEVATION_FAILED
                Exit Sub
            Case Else
                sCmdLine = sCmdLine  &  " " & chr(34) & Argument & chr(34)
            End Select
        Next 'Argument
    End If
    ' prep work to get the return value from the elevated process
    iParentProcessId = GetMyProcessId 
    
'    ' make user aware of elevation attempt after reboot
'    If RegReadValue(HKCU, "SOFTWARE\Microsoft\Office\15.0\CleanC2R", "Rerun", sValue, "REG_DWORD") Then
'        oWShell.Popup "System reboot complete. OffScrub will now prompt for elevation!", 10, SCRIPTNAME & " - NOTE!"
'    End If
    
    ' launch the elevated instance
    oShell.ShellExecute "cscript.exe", sCmdLine & " /NoElevate UAC", "", "runas", 1
    ' get the process id of the spawned instance
    WScript.Sleep 500
    Set Processes = oWmiLocal.ExecQuery("Select * From Win32_Process WHERE ParentProcessId='" & iParentProcessId & "'")
    If Processes.Count > 0 Then
        For Each Process in Processes
		    iSpawnedProcessId = Process.ProcessId
		    Exit For
        Next 'Process
        ' monitor the tasklist to detect the end of the spawned process
        While oWmiLocal.ExecQuery("Select * From Win32_Process WHERE ProcessId='" & iSpawnedProcessId & "'").Count > 0
            WScript.Sleep 3000
        Wend
        ' get the return value from the file
        Wscript.Quit GetRetValFromFile
    End If
    ' elevation failed (user declined)
    SetError ERROR_ELEVATION_USERDECLINED
End Sub 'RelaunchElevated

'-------------------------------------------------------------------------------
'   RelaunchAsCScript
'
'   Relaunch self with Cscript as host
'-------------------------------------------------------------------------------
Sub RelaunchAsCScript
    Dim Argument
    Dim sCmdLine
    Dim fQuietNoCScript

    fQuietNoCScript = False
    SetError ERROR_RELAUNCH
    sCmdLine = "cmd.exe /c " & WScript.Path & "\cscript.exe //NOLOGO " & Chr(34) & WScript.scriptFullName & Chr(34)
    If Wscript.Arguments.Count > 0 Then
        For Each Argument in Wscript.Arguments
            sCmdLine = sCmdLine  &  " " & chr(34) & Argument & chr(34)
            Select Case UCase(Argument)
            Case "/Q","/QUIET"
                fQuietNoCScript = True
                ClearError ERROR_RELAUNCH
            End Select
        Next 'Argument
    End If
    sCmdLine = sCmdLine  &  " " & chr(34) & "/ChangedScriptHost" & chr(34)
    
    If NOT fQuietNoCScript Then Wscript.Quit CLng(oWShell.Run(sCmdLine, 1, True))
End Sub 'RelaunchAsCScript

'-------------------------------------------------------------------------------
'   SetError
'
'   Set error bit(s) 
'-------------------------------------------------------------------------------
Sub SetError(ErrorBit)
    iError = iError OR ErrorBit
    Select Case ErrorBit
    Case ERROR_DCAF_FAILURE, ERROR_STAGE2, ERROR_ELEVATION_USERDECLINED, ERROR_ELEVATION, ERROR_SCRIPTINIT
        iError = iError OR ERROR_FAIL
    End Select
End Sub

'-------------------------------------------------------------------------------
'   ClearError
'
'   Unset error bit(s) 
'-------------------------------------------------------------------------------
Sub ClearError(ErrorBit)
    iError = iError AND (ERROR_ALL - ErrorBit)
    Select Case ErrorBit
    Case ERROR_ELEVATION_USERDECLINED, ERROR_ELEVATION, ERROR_SCRIPTINIT
        iError = iError AND (ERROR_ALL - ERROR_FAIL)
    End Select
End Sub

'-------------------------------------------------------------------------------
'   SetRetVal
'
'   Write return value to file
'-------------------------------------------------------------------------------
Sub SetRetVal(iError)
    Dim RetValFileStream
    
    'don't fail script execution if writing the return value to file fails
    On Error Resume Next 

    Set RetValFileStream = oFso.createTextFile(sScrubDir & "\" & RETVALFILE, True, True)
    RetValFileStream.Write iError
    RetValFileStream.Close
    On Error Goto 0
End Sub 'SetRetVal

'-------------------------------------------------------------------------------
'   GetRetValFromFile
'
'   Read return value from file.
'   Used to ensure return value can get obtained from an elevated process
'-------------------------------------------------------------------------------
Function GetRetValFromFile ()
    Dim RetValFileStream
    Dim iRetValFromFile

    On Error Resume Next 'don't fail script execution when getting the return value from file fails

    If oFso.FileExists(sScrubDir & "\" & RETVALFILE) Then
        Set RetValFileStream = oFso.OpenTextFile(sScrubDir & "\" & RETVALFILE, 1, False, -2)
        GetRetValFromFile = RetValFileStream.ReadAll
        RetValFileStream.Close
        Exit Function
    End If
    Err.Clear
    On Error Goto 0
    GetRetValFromFile = ERROR_UNKNOWN
End Function 'GetRetValFromFile

'-------------------------------------------------------------------------------
'   CreateLog
'
'   Create the removal log file
'-------------------------------------------------------------------------------
Sub CreateLog
    Dim DateTime
    Dim sLogName
    
    On Error Resume Next
    ' create the log file
    Set DateTime = CreateObject("WbemScripting.SWbemDateTime")
    DateTime.SetVarDate Now, True
    sLogName = sLogDir & "\" & oWShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
    sLogName = sLogName &  "_" & Left(DateTime.Value, 14)
    sLogName = sLogName & "_ScrubLog.txt"
    Err.Clear
    Set LogStream = oFso.CreateTextFile(sLogName, True, True)
    If Err <> 0 Then 
        Err.Clear
        sLogDir = sScrubDir
        sLogName = sLogDir & "\" & oWShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
        sLogName = sLogName &  "_" & Left(DateTime.Value, 14)
        sLogName = sLogName & "_ScrubLog.txt"
        Set LogStream = oFso.CreateTextFile(sLogName, True, True)
    End If
    On Error Goto 0

    LogH2 "Microsoft Customer Support Services - " & ONAME & " Removal Utility" & vbCrLf & vbCrLf & _
        	"Version: " & vbTab & SCRIPTVERSION & vbCrLf & _
        	"64 bit OS: " & vbTab & f64 & vbCrLf & _
        	"Removal start: " & vbTab & Time 
    LogH2	"OS Details: " & sOSinfo & vbCrLf
    fLogInitialized = True
End Sub 'CreateLog

'-------------------------------------------------------------------------------
'   HiveString
'
'   Translates the numeric constant into the human readable registry hive string
'-------------------------------------------------------------------------------
Function HiveString(hDefKey)
    Select Case hDefKey
        Case HKCR : HiveString = "HKEY_CLASSES_ROOT"
        Case HKCU : HiveString = "HKEY_CURRENT_USER"
        Case HKLM : HiveString = "HKEY_LOCAL_MACHINE"
        Case HKU  : HiveString = "HKEY_USERS"
        Case Else : HiveString = hDefKey
    End Select
End Function

'-------------------------------------------------------------------------------
'   RegKeyExists
'
'   Returns a boolean for the test on existence of a given registry key
'-------------------------------------------------------------------------------
Function RegKeyExists(hDefKey, sSubKeyName)
    Dim arrKeys
    RegKeyExists = False
    If oReg.EnumKey(hDefKey, sSubKeyName, arrKeys) = 0 Then RegKeyExists = True
End Function

'-------------------------------------------------------------------------------
'   RegValExists
'
'   Returns a boolean for the test on existence of a given registry value
'-------------------------------------------------------------------------------
Function RegValExists(hDefKey,sSubKeyName,sName)
    Dim arrValueTypes, arrValueNames
    Dim i

    RegValExists = False
    If Not RegKeyExists(hDefKey,sSubKeyName) Then Exit Function
    If oReg.EnumValues(hDefKey,sSubKeyName,arrValueNames,arrValueTypes) = 0 AND IsArray(arrValueNames) Then
        For i = 0 To UBound(arrValueNames) 
            If LCase(arrValueNames(i)) = Trim(LCase(sName)) Then RegValExists = True
        Next 
    End If 'oReg.EnumValues
End Function

'-------------------------------------------------------------------------------
'   RegReadValue
'
'   Read the value of a given registry entry
'   The correct type has to be passed in as argument
'-------------------------------------------------------------------------------
Function RegReadValue(hDefKey, sSubKeyName, sName, sValue, sType)
    Dim RetVal
    Dim Item
    Dim arrValues
    
    Select Case UCase(sType)
        Case "1", "REG_SZ"
            RetVal = oReg.GetStringValue(hDefKey, sSubKeyName, sName, sValue)
            If Not RetVal = 0 AND f64 Then RetVal = oReg.GetStringValue(hDefKey, Wow64Key(hDefKey, sSubKeyName), sName, sValue)
        Case "2", "REG_EXPAND_SZ"
            RetVal = oReg.GetExpandedStringValue(hDefKey, sSubKeyName, sName, sValue)
            If NOT RetVal = 0 AND f64 Then RetVal = oReg.GetExpandedStringValue(hDefKey, Wow64Key(hDefKey, sSubKeyName), sName, sValue)
        Case "3", "REG_BINARY"
            RetVal = oReg.GetBinaryValue(hDefKey, sSubKeyName, sName, sValue)
            If NOT RetVal = 0 AND f64 Then RetVal = oReg.GetBinaryValue(hDefKey, Wow64Key(hDefKey, sSubKeyName), sName, sValue)
        Case "4", "REG_DWORD"
            RetVal = oReg.GetDWORDValue(hDefKey, sSubKeyName, sName, sValue)
            If NOT RetVal = 0 AND f64 Then RetVal = oReg.GetDWORDValue(hDefKey, Wow64Key(hDefKey, sSubKeyName), sName, sValue)
        Case "7", "REG_MULTI_SZ"
            RetVal = oReg.GetMultiStringValue(hDefKey, sSubKeyName, sName, arrValues)
            If NOT RetVal = 0 AND f64 Then RetVal = oReg.GetMultiStringValue(hDefKey, Wow64Key(hDefKey, sSubKeyName), sName, arrValues)
            If RetVal = 0 Then sValue = Join(arrValues, chr(13))
        Case Else
            RetVal = -1
    End Select 'sValue
    
    RegReadValue = (RetVal = 0)
End Function 'RegReadValue

'-------------------------------------------------------------------------------
'   RegEnumValues
'
'   Enumerate a registry key to return all values
'-------------------------------------------------------------------------------
Function RegEnumValues(hDefKey, sSubKeyName, arrNames, arrTypes)
    Dim RetVal, RetVal64
    Dim arrNames32, arrNames64, arrTypes32, arrTypes64
    
    If f64 Then
        RetVal = oReg.EnumValues(hDefKey, sSubKeyName, arrNames32, arrTypes32)
        RetVal64 = oReg.EnumValues(hDefKey, Wow64Key(hDefKey, sSubKeyName), arrNames64, arrTypes64)
        If (RetVal = 0) AND (NOT RetVal64 = 0) AND IsArray(arrNames32) AND IsArray(arrTypes32) Then 
            arrNames = arrNames32
            arrTypes = arrTypes32
        End If
        If (NOT RetVal = 0) AND (RetVal64 = 0) AND IsArray(arrNames64) AND IsArray(arrTypes64) Then 
            arrNames = arrNames64
            arrTypes = arrTypes64
        End If
        If (RetVal = 0) AND (RetVal64 = 0) AND IsArray(arrNames32) AND IsArray(arrNames64) AND IsArray(arrTypes32) AND IsArray(arrTypes64) Then 
            arrNames = RemoveDuplicates(Split((Join(arrNames32, "\") & "\" & Join(arrNames64, "\")), "\"))
            arrTypes = RemoveDuplicates(Split((Join(arrTypes32, "\") & "\" & Join(arrTypes64, "\")), "\"))
        End If
    Else
        RetVal = oReg.EnumValues(hDefKey, sSubKeyName, arrNames, arrTypes)
    End If 'f64
    RegEnumValues = ((RetVal = 0) OR (RetVal64 = 0)) AND IsArray(arrNames) AND IsArray(arrTypes)
End Function 'RegEnumValues

'-------------------------------------------------------------------------------
'   RegEnumKey
'
'   Enumerate a registry key to return all subkeys
'-------------------------------------------------------------------------------
Function RegEnumKey(hDefKey, sSubKeyName, arrKeys)
    Dim RetVal, RetVal64
    Dim arrKeys32, arrKeys64
    
    If f64 Then
        RetVal = oReg.EnumKey(hDefKey, sSubKeyName, arrKeys32)
        RetVal64 = oReg.EnumKey(hDefKey, Wow64Key(hDefKey, sSubKeyName), arrKeys64)
        If (RetVal = 0) AND (NOT RetVal64 = 0) AND IsArray(arrKeys32) Then arrKeys = arrKeys32
        If (Not RetVal = 0) AND (RetVal64 = 0) AND IsArray(arrKeys64) Then arrKeys = arrKeys64
        If (RetVal = 0) AND (RetVal64 = 0) Then 
            If IsArray(arrKeys32) AND IsArray (arrKeys64) Then 
                arrKeys = RemoveDuplicates(Split((Join(arrKeys32, "\") & "\" & Join(arrKeys64, "\")), "\"))
            ElseIf IsArray(arrKeys64) Then
                arrKeys = arrKeys64
            Else
                arrKeys = arrKeys32
            End If
        End If
    Else
        RetVal = oReg.EnumKey(hDefKey, sSubKeyName, arrKeys)
    End If 'f64
    RegEnumKey = ((RetVal = 0) OR (RetVal64 = 0)) AND IsArray(arrKeys)
End Function 'RegEnumKey

'-------------------------------------------------------------------------------
'   RegDeleteValue
'
'   Wrapper around oReg.DeleteValue to handle 64 bit
'-------------------------------------------------------------------------------
Sub RegDeleteValue(hDefKey, sSubKeyName, sName, fRegMultiSZ)
    Dim sDelKeyName, sValue
    Dim iRetVal
    Dim fKeep
    
    ' ensure trailing "\"
    sSubKeyName = sSubKeyName & "\"
    While InStr(sSubKeyName, "\\") > 0
        sSubKeyName = Replace(sSubKeyName, "\\", "\")
    Wend

    fKeep = dicKeepReg.Exists(LCase(sSubKeyName & sName))
    If (NOT fKeep AND f64) Then fKeep = dicKeepReg.Exists(LCase(Wow64Key(hDefKey, sSubKeyName) & sName))
    If fKeep Then
        LogOnly "Disallowing the delete of still required keypath element: " & HiveString(hDefKey) & "\" & sSubKeyName & sName
        If NOT fForce Then Exit Sub
    End If
    
    ' check on forced delete
    If fKeep Then
        LogOnly "Enforced delete of still required keypath element: " & HiveString(hDefKey) & "\" & sSubKeyName & sName
        LogOnly "   Remaining applications will need a repair!"
    End If
    
    ' ensure value exists
    If RegValExists(hDefKey, sSubKeyName, sName) Then
        sDelKeyName = sSubKeyName
    ElseIf RegValExists(hDefKey, Wow64Key(hDefKey, sSubKeyName), sName) Then
        sDelKeyName =  Wow64Key(hDefKey, sSubKeyName)
    Else
        LogOnly "Value not found. Cannot delete value: " & HiveString(hDefKey) & "\" & sSubKeyName & sName
        Exit Sub
    End If

    ' prevent unintentional, unsafe REG_MULTI_SZ delete
    If RegReadValue(hDefKey, sDelKeyName, sName, sValue, "REG_MULTI_SZ") AND NOT fRegMultiSZ Then
        LogOnly "Disallowing unsafe delete of REG_MULTI_SZ: " & HiveString(hDefKey) & "\" & sDelKeyName & sName
        Exit Sub
    End If
    
    ' execute delete operation
    If Not fDetectOnly Then 
        LogOnly "Delete registry value: " & HiveString(hDefKey) & "\" & sDelKeyName & " -> " & sName
        iRetVal = 0
        iRetVal = oReg.DeleteValue(hDefKey, sDelKeyName, sName)
        CheckError "RegDeleteValue"
        If NOT (iRetVal = 0) Then
            LogOnly "     Delete failed. Return value: " & iRetVal
            SetError ERROR_STAGE2
        End If
    Else
        LogOnly "Preview mode. Disallowing delete registry value: " & HiveString(hDefKey) & "\" & sDelKeyName & " -> " & sName
    End If
    On Error Goto 0

End Sub 'RegDeleteValue

'-------------------------------------------------------------------------------
'   RegDeleteKey
'
'   Wrappper around RegDeleteKeyEx to handle 64bit
'-------------------------------------------------------------------------------
Sub RegDeleteKey(hDefKey, sSubKeyName)
    Dim sDelKeyName
    Dim fKeep
    
    ' ensure trailing "\"
    sSubKeyName = sSubKeyName & "\"
    While InStr(sSubKeyName, "\\") > 0
        sSubKeyName = Replace(sSubKeyName, "\\", "\")
    Wend

    fKeep = dicKeepReg.Exists(LCase(sSubKeyName))
    If (NOT fKeep AND f64) Then fKeep = dicKeepReg.Exists(LCase(Wow64Key(hDefKey, sSubKeyName)))
    If fKeep Then
        LogOnly "Disallowing the delete of still required keypath element: " & HiveString(hDefKey) & "\" & sSubKeyName
        If NOT fForce Then Exit Sub
    End If
    
    ' check on forced delete
    If fKeep Then
        LogOnly "Enforced delete of still required keypath element: " & HiveString(hDefKey) & "\" & sSubKeyName
        LogOnly "   Remaining applications will need a repair!"
    End If
    
    If Len(sSubKeyName) > 1 Then
        'Strip of trailing "\"
        sSubKeyName = Left(sSubKeyName, Len(sSubKeyName) - 1)
    End If
    
    ' ensure key exists
    If RegKeyExists(hDefKey, sSubKeyName) Then
        sDelKeyName = sSubKeyName
    ElseIf f64 AND RegKeyExists(hDefKey, Wow64Key(hDefKey, sSubKeyName)) Then
        sDelKeyName = Wow64Key(hDefKey, sSubKeyName)
    Else
        LogOnly "Key not found. Cannot delete key: " & HiveString(hDefKey) & "\" & sSubKeyName
        Exit Sub
    End If

    ' execute delete
    If Not fDetectOnly Then
        LogOnly "Delete registry key: " & HiveString(hDefKey) & "\" & sDelKeyName
        On Error Resume Next
        RegDeleteKeyEx hDefKey, sDelKeyName
        On Error Goto 0
    Else
        LogOnly "Preview mode. Disallowing delete of registry key: " & HiveString(hDefKey) & "\" & sSubKeyName
    End If
End Sub 'RegDeleteKey

'-------------------------------------------------------------------------------
'   RegDeleteKeyEx
'
'   Recursively delete a registry structure
'-------------------------------------------------------------------------------
Sub RegDeleteKeyEx(hDefKey, sSubKeyName) 
    Dim arrSubkeys
    Dim sSubkey
    Dim iRetVal

    'Strip of trailing "\"
    If Len(sSubKeyName) > 1 Then
        If Right(sSubKeyName, 1) = "\" Then sSubKeyName = Left(sSubKeyName, Len(sSubKeyName) - 1)
    End If
    On Error Resume Next

    ' exception handler
    If (hDefKey = HKLM) AND (sSubKeyName = "SOFTWARE\Microsoft\Office\15.0\ClickToRun") Then
        iRetVal = oWShell.Run("reg delete HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun /f", 0, True)
        Exit Sub
    End If

    ' regular recursion
    oReg.EnumKey hDefKey, sSubKeyName, arrSubkeys
    If IsArray(arrSubkeys) Then 
        For Each sSubkey In arrSubkeys 
            RegDeleteKeyEx hDefKey, sSubKeyName & "\" & sSubkey 
        Next 
    End If 
    If Not fDetectOnly Then 
        iRetVal = 0
        iRetVal = oReg.DeleteKey(hDefKey, sSubKeyName)
        If NOT (iRetVal = 0) Then LogOnly "     Delete failed. Return value: "&iRetVal
    End If
    On Error Goto 0
End Sub 'RegDeleteKeyEx

'-------------------------------------------------------------------------------
'   Wow64Key
'
'   Return the 32bit regkey location on a 64bit environment
'-------------------------------------------------------------------------------
Function Wow64Key(hDefKey, sSubKeyName)
    Dim iPos

    Select Case hDefKey
    Case HKCU
        If Left(sSubKeyName, 17) = "Software\Classes\" Then
            Wow64Key = Left(sSubKeyName, 17) & "Wow6432Node\" & Right(sSubKeyName, Len(sSubKeyName) - 17)
        Else
            iPos = InStr(sSubKeyName, "\")
            Wow64Key = Left(sSubKeyName, iPos) & "Wow6432Node\" & Right(sSubKeyName, Len(sSubKeyName) - iPos)
        End If
    Case HKLM
        If Left(sSubKeyName, 17) = "Software\Classes\" Then
            Wow64Key = Left(sSubKeyName, 17) & "Wow6432Node\" & Right(sSubKeyName, Len(sSubKeyName) - 17)
        Else
            iPos = InStr(sSubKeyName, "\")
            Wow64Key = Left(sSubKeyName, iPos) & "Wow6432Node\" & Right(sSubKeyName, Len(sSubKeyName) - iPos)
        End If
    Case Else
        Wow64Key = "Wow6432Node\" & sSubKeyName
    End Select 'hDefKey
End Function 'Wow64Key

'-------------------------------------------------------------------------------
'   RemoveDuplicates
'
'   Remove duplicate entries from a one dimensional array
'-------------------------------------------------------------------------------
Function RemoveDuplicates(Array)
    Dim Item
    Dim dicNoDupes
    
    Set dicNoDupes = CreateObject("Scripting.Dictionary")
    For Each Item in Array
        If Not dicNoDupes.Exists(Item) Then dicNoDupes.Add Item,Item
    Next 'Item
    RemoveDuplicates = dicNoDupes.Keys
End Function 'RemoveDuplicates

'-------------------------------------------------------------------------------
'   CheckError
'
'   Checks the status of 'Err' and logs the error details if <> 0
'-------------------------------------------------------------------------------
Sub CheckError(sModule)
    If Err <> 0 Then 
        LogOnly "   Error: " & sModule & " - Source: " & Err.Source & "; Err# (Hex): " & Hex( Err ) & _
               "; Err# (Dec): " & Err & "; Description : " & Err.Description
    End If 'Err = 0
    Err.Clear
End Sub

'-------------------------------------------------------------------------------
'   LogH
'
'   Write a header log string to the log file
'-------------------------------------------------------------------------------
Sub LogH (sLog)
    LogStream.WriteLine ""
    sLog = sLog & vbCrLf & String(Len(sLog), "=")
    If NOT fQuiet AND fCScript Then wscript.echo ""
    If NOT fQuiet AND fCScript Then wscript.echo sLog
    LogStream.WriteLine sLog
End Sub 'Logh

'-------------------------------------------------------------------------------
'   LogH1
'
'   Write a header log string to the log file
'-------------------------------------------------------------------------------
Sub LogH1 (sLog)
    LogStream.WriteLine ""
    sLog = sLog & vbCrLf & String(Len(sLog), "-")
    If NOT fQuiet AND fCScript Then wscript.echo ""
    If NOT fQuiet AND fCScript Then wscript.echo sLog
    LogStream.WriteLine sLog
End Sub 'LogH1

'-------------------------------------------------------------------------------
'   LogH2
'
'   Write w/o indent Cmd window and the log file
'-------------------------------------------------------------------------------
Sub LogH2 (sLog)
    If NOT fQuiet AND fCScript Then wscript.echo sLog
    LogStream.WriteLine ""
    LogStream.WriteLine sLog
End Sub 'LogH2

'-------------------------------------------------------------------------------
'   Log
'
'   Echos the log string to the Cmd window and the log file
'-------------------------------------------------------------------------------
Sub Log (sLog)
    If NOT fQuiet AND fCScript Then wscript.echo sLog
    If sLog = "" Then
        LogStream.WriteLine
    Else
        LogStream.WriteLine "   " & Time & ": " & sLog
    End If
End Sub 'Log

'-------------------------------------------------------------------------------
'   LogOnly
'
'   Commits the log string to the log file
'-------------------------------------------------------------------------------
Sub LogOnly (sLog)
    If sLog = "" Then
        LogStream.WriteLine
    Else
        LogStream.WriteLine "   " & Time & ": " & sLog
    End If
End Sub 'Log

'-------------------------------------------------------------------------------
'   InScope
'
'   Check if ProductCode is in scope for removal
'-------------------------------------------------------------------------------
'Check if ProductCode is in scope
Function InScope(sProductCode)
    Dim fInScope
    Dim sProd
	
	Const OFFICEID = "0000000FF1CE}"
	
	On Error Resume Next

    fInScope = False
    'LogOnly "Now checking scope of: " & sProductCode
    If Len(sProductCode) = 38 Then
        'LogOnly "GUID length validated to be 38 characters"
        sProd = UCase(sProductCode)
        If Right(sProd, PRODLEN) = OFFICEID Then
        	'LogOnly "Pattern matches " & OFFICEID
        	If CInt(Mid(sProd, 4, 2)) > 14 Then 
	            If Err <> 0 Then
	            	Err.Clear
	            	Exit Function
	            End If
	            'LogOnly "VersionMajor confirmed to be > 14" 
	            Select Case Mid(sProd, 11, 4)
	            Case "007E", "008F", "008C", "24E1", "237A", "00DD"
	                'LogOnly "SKUFilter matches scope"
	                fInScope = True
	            Case Else
	            	'LogOnly "SKU " & Mid(sProd, 11, 4) & " doesn't match known integration products scope"
	            End Select
            End If
        End If
        ' Microsoft Online Services Sign-in Assistant (x64 ship and x86 ship)
        If sProd = "{6C1ADE97-24E1-4AE4-AEDD-86D3A209CE60}" Then fInScope = True
        If sProd = "{9520DDEB-237A-41DB-AA20-F2EF2360DCEB}" Then fInScope = True
        If sProd = UCase(sPackageGuid) Then fInScope = True
        If sProd = UCase("{9AC08E99-230B-47e8-9721-4577B7F124EA}") Then fInScope = True
    End If '38

    InScope = fInScope
End Function 'InScope

'-------------------------------------------------------------------------------
'   CheckDelete
'
'   Check a ProductCode is known to stay installed
'-------------------------------------------------------------------------------
Function CheckDelete(sProductCode)

    CheckDelete = False
    ' ensure valid GUID length
    If NOT Len(sProductCode) = 38 Then Exit Function
    ' only care if it's in the expected ProductCode pattern 
    If NOT InScope(sProductCode) Then Exit Function
    ' check if it's a known product that should be kept
    If dicKeepSku.Exists(UCase(sProductCode)) Then Exit Function
    
    CheckDelete = True
End Function 'CheckDelete

'-------------------------------------------------------------------------------
'   DeleteService
'
'   Delete a service
'-------------------------------------------------------------------------------
'Delete a service
Sub DeleteService(sName)
    Dim Services, srvc, Processes, process
    Dim sQuery, sStates, sProcessName, sCmd
    Dim iRet
    
    On Error Resume Next
    
    sStates = "STARTED;RUNNING"
    sQuery = "Select * From Win32_Service Where Name='" & sName & "'"
    Set Services = oWmiLocal.Execquery(sQuery)
    
    ' stop and delete the service
    For Each srvc in Services
        Log "   Found service " & sName & " (" & srvc.DisplayName & ") in state " & srvc.State
        ' get the process name
        sProcessName = Trim(Replace(Mid(srvc.PathName, InStrRev(srvc.PathName,"\") + 1), chr(34), ""))
        ' stop the service
        If InStr(sStates, UCase(srvc.State)) > 0 Then
            iRet = srvc.StopService()
            LogOnly " attempt to stop service " & sName & " returned: " & iRet
        End If
        ' ensure no more instances of the service are running
        Set Processes = oWmiLocal.ExecQuery("Select * From Win32_Process Where Name='" & sProcessName & "'")
        For Each process in Processes
            iRet = process.Terminate()
        Next 'Process
        If fDetectOnly Then 
            Log "   Not deleting service " & sName & " in preview mode"
            Exit Sub
        End If
        iRet = srvc.Delete()
        Log "   Delete service " & sName & " returned: " & iRet
    Next 'srvc

    ' check if service got deleted
    Set Services = oWmiLocal.Execquery(sQuery)
    For Each srvc in Services
        ' failed to delete service. retry with 'sc' command
        sLog "Delete service " & sName & " failed."
        sLog "Retry delete using 'SC' command"
        sCmd = "sc delete " & sName
        iRet = oWShell.Run(sCmd, 0, True)
    Next 'srvc

    Set Services = Nothing
    Err.Clear
    On Error Goto 0
End Sub 'DeleteService


'-------------------------------------------------------------------------------
'   SetupRetVal
'
'   Translation for known uninstall return values
'-------------------------------------------------------------------------------
Function SetupRetVal(RetVal)
    Select Case RetVal
        Case 0 : SetupRetVal = "Success"
        'msiexec return values
        Case 1259 : SetupRetVal = "APPHELP_BLOCK"
        Case 1601 : SetupRetVal = "INSTALL_SERVICE_FAILURE"
        Case 1602 : SetupRetVal = "INSTALL_USEREXIT"
        Case 1603 : SetupRetVal = "INSTALL_FAILURE"
        Case 1604 : SetupRetVal = "INSTALL_SUSPEND"
        Case 1605 : SetupRetVal = "UNKNOWN_PRODUCT"
        Case 1606 : SetupRetVal = "UNKNOWN_FEATURE"
        Case 1607 : SetupRetVal = "UNKNOWN_COMPONENT"
        Case 1608 : SetupRetVal = "UNKNOWN_PROPERTY"
        Case 1609 : SetupRetVal = "INVALID_HANDLE_STATE"
        Case 1610 : SetupRetVal = "BAD_CONFIGURATION"
        Case 1611 : SetupRetVal = "INDEX_ABSENT"
        Case 1612 : SetupRetVal = "INSTALL_SOURCE_ABSENT"
        Case 1613 : SetupRetVal = "INSTALL_PACKAGE_VERSION"
        Case 1614 : SetupRetVal = "PRODUCT_UNINSTALLED"
        Case 1615 : SetupRetVal = "BAD_QUERY_SYNTAX"
        Case 1616 : SetupRetVal = "INVALID_FIELD"
        Case 1618 : SetupRetVal = "INSTALL_ALREADY_RUNNING"
        Case 1619 : SetupRetVal = "INSTALL_PACKAGE_OPEN_FAILED"
        Case 1620 : SetupRetVal = "INSTALL_PACKAGE_INVALID"
        Case 1621 : SetupRetVal = "INSTALL_UI_FAILURE"
        Case 1622 : SetupRetVal = "INSTALL_LOG_FAILURE"
        Case 1623 : SetupRetVal = "INSTALL_LANGUAGE_UNSUPPORTED"
        Case 1624 : SetupRetVal = "INSTALL_TRANSFORM_FAILURE"
        Case 1625 : SetupRetVal = "INSTALL_PACKAGE_REJECTED"
        Case 1626 : SetupRetVal = "FUNCTION_NOT_CALLED"
        Case 1627 : SetupRetVal = "FUNCTION_FAILED"
        Case 1628 : SetupRetVal = "INVALID_TABLE"
        Case 1629 : SetupRetVal = "DATATYPE_MISMATCH"
        Case 1630 : SetupRetVal = "UNSUPPORTED_TYPE"
        Case 1631 : SetupRetVal = "CREATE_FAILED"
        Case 1632 : SetupRetVal = "INSTALL_TEMP_UNWRITABLE"
        Case 1633 : SetupRetVal = "INSTALL_PLATFORM_UNSUPPORTED"
        Case 1634 : SetupRetVal = "INSTALL_NOTUSED"
        Case 1635 : SetupRetVal = "PATCH_PACKAGE_OPEN_FAILED"
        Case 1636 : SetupRetVal = "PATCH_PACKAGE_INVALID"
        Case 1637 : SetupRetVal = "PATCH_PACKAGE_UNSUPPORTED"
        Case 1638 : SetupRetVal = "PRODUCT_VERSION"
        Case 1639 : SetupRetVal = "INVALID_COMMAND_LINE"
        Case 1640 : SetupRetVal = "INSTALL_REMOTE_DISALLOWED"
        Case 1641 : SetupRetVal = "SUCCESS_REBOOT_INITIATED"
        Case 1642 : SetupRetVal = "PATCH_TARGET_NOT_FOUND"
        Case 1643 : SetupRetVal = "PATCH_PACKAGE_REJECTED"
        Case 1644 : SetupRetVal = "INSTALL_TRANSFORM_REJECTED"
        Case 1645 : SetupRetVal = "INSTALL_REMOTE_PROHIBITED"
        Case 1646 : SetupRetVal = "PATCH_REMOVAL_UNSUPPORTED"
        Case 1647 : SetupRetVal = "UNKNOWN_PATCH"
        Case 1648 : SetupRetVal = "PATCH_NO_SEQUENCE"
        Case 1649 : SetupRetVal = "PATCH_REMOVAL_DISALLOWED"
        Case 1650 : SetupRetVal = "INVALID_PATCH_XML"
        Case 3010 : SetupRetVal = "SUCCESS_REBOOT_REQUIRED"
        Case Else : SetupRetVal = "Unknown Return Value"
    End Select
End Function 'SetupRetVal

'-------------------------------------------------------------------------------
'   DeleteFile
'
'   Wrapper to delete a file
'-------------------------------------------------------------------------------
Sub DeleteFile(sFile)
    Dim File, attr
    Dim sDelFile, sFileName, sNewPath
    Dim fKeep
    
    On Error Resume Next

    fKeep = dicKeepFolder.Exists(LCase(sFile))
    If (NOT fKeep AND f64) Then fKeep = dicKeepFolder.Exists(LCase(Wow64Folder(sFile)))
    If fKeep Then
        LogOnly "Disallowing the delete of still required keypath element: " & sFile
        If NOT fForce Then Exit Sub
    End If

    ' check on forced delete
    If fKeep Then
        LogOnly "Enforced delete of still required keypath element: " & sFile
        LogOnly "   Remaining applications will need a repair!"
    End If

    If oFso.FileExists(sFile) Then
        sDelFile = sFile
    ElseIf f64 AND oFso.FileExists(Wow64Folder(sFile)) Then
        sDelFile = Wow64Folder(sFile)
    Else
        LogOnly "Path not found. Cannot not delete folder: " & sFile
        Exit Sub
    End If
    If Not fDetectOnly Then 
        LogOnly "Delete file: " & sDelFile
		Set File = oFso.GetFile(sDelFile)
        ' ensure read-only flag is not set
        attr =  File.Attributes
        If CBool(attr AND 1) Then File.Attributes = attr AND (attr - 1)
        ' add folder to empty folder cleanup list
        If NOT dicDelFolder.Exists(File.ParentFolder.Path) Then dicDelFolder.Add File.ParentFolder.Path, File.ParentFolder.Path
        ' delete the file
        sFile = File.Path
        File.Delete True
        Set File = Nothing
        If Err <> 0 Then
            CheckError "DeleteFile"
            ' schedule file for delete on next reboot
            ScheduleDeleteFile sFile
        End If 'Err <> 0
    Else
        LogOnly "Preview mode. Disallowing delete for folder: " & sDelFile
    End If
    On Error Goto 0
End Sub 'DeleteFile

'-------------------------------------------------------------------------------
'   DeleteFolder
'
'   Wrapper to delete a folder
'-------------------------------------------------------------------------------
Sub DeleteFolder(sFolder)
    Dim Folder, fld, attr
    Dim sDelFolder, sFolderName, sNewPath, sCmd
    Dim fKeep
    
    ' ensure trailing "\"
    ' trailing \ is required for dicKeepFolder comparisons
    sFolder = sFolder & "\"
    While InStr(sFolder,"\\")>0
        sFolder = Replace(sFolder,"\\","\")
    Wend

    ' prevent delete of folders that are known to be still required
    fKeep = dicKeepFolder.Exists(LCase(sFolder))
    If (NOT fKeep AND f64) Then fKeep = dicKeepFolder.Exists(LCase(Wow64Folder(sFolder)))
    If fKeep Then
        LogOnly "Disallowing the delete of still required keypath element: " & sFolder
        If NOT fForce Then Exit Sub
    End If

    ' check on forced delete
    If fKeep Then
        LogOnly "Enforced delete of still required keypath element: " & sFolder
        LogOnly "   Remaining applications will need a repair!"
    End If
    
    ' strip trailing "\"
    If Len(sFolder) > 1 Then
        sFolder = Left(sFolder, Len(sFolder) - 1)
    End If

    On Error Resume Next
    If oFso.FolderExists(sFolder) Then 
        sDelFolder = sFolder
    ElseIf f64 AND oFso.FolderExists(Wow64Folder(sFolder)) Then 
        sDelFolder = Wow64Folder(sFolder)
    Else
        LogOnly "Path not found. Cannot not delete folder: " & sFolder
        Exit Sub
    End If
    If Not fDetectOnly Then 
        LogOnly "Delete folder: " & sDelFolder
        Set Folder = oFso.GetFolder(sDelFolder)
        ' ensure to remove read only flag
        attr =  Folder.Attributes
        If CBool(attr AND 1) Then Folder.Attributes = attr AND (attr - 1)
        ' add to empty folder cleanup list
        If NOT dicDelFolder.Exists(Folder.Path) Then dicDelFolder.Add Folder.Path, Folder.Path
        ' delete the folder
        ' for performance reasons try 'rd' first
        Set Folder = Nothing
        sCmd = "cmd.exe /c rd /s " & chr(34) & sDelFolder & chr(34) & " /q"
        oWShell.Run sCmd, 0, True
        If NOT oFso.FolderExists(sDelFolder) Then Exit Sub
        
        ' rd didn't work check with FileSystemObject
        Set Folder = oFso.GetFolder(sDelFolder)
        Folder.Delete True
        Set Folder = Nothing
        
        ' error handling
        If Err <> 0 Then
            Select Case Err
            Case 70
                ' Access Denied
                ' Retry after closing running processes
                CheckError "DeleteFolder"
                If NOT fRerun Then
                    CloseOfficeApps
                    ' attempt 'rd' command
                    LogOnly "   Attempt to remove with 'rd' command"
                    sCmd = "cmd.exe /c rd /s " & chr(34) & sDelFolder & chr(34) & " /q"
                    oWShell.Run sCmd, 0, True
                    If NOT oFso.FolderExists(sDelFolder) Then Exit Sub
                End If

            Case 76 
                ' check on invalid path lengt issues Err 76 (0x4C) "Path not found"
                ' attempt 'rd' command
                CheckError "DeleteFolder"
                LogOnly "   Attempt to remove with 'rd' command"
                sCmd = "cmd.exe /c rd /s " & chr(34) & sDelFolder & chr(34) & " /q"
                oWShell.Run sCmd, 0, True
                If NOT oFso.FolderExists(sDelFolder) Then Exit Sub
            End Select
            
            ' stil failed!
            Log "   Failed to delete folder: " & sDelFolder
            CheckError "DeleteFolder"

            ' try to delete as many folder contents as possible
            ' before the recursive error handling is called
            Set Folder = oFso.GetFolder(sDelFolder)
            For Each fld in Folder.Subfolders
                sCmd = "cmd.exe /c rd /s " & chr(34) & fld.Path & chr(34) & " /q"
                oWShell.Run sCmd, 0, True
            Next 'fld
            sCmd = "cmd.exe /c del " & chr(34) & fld.Path & "\*.*" & chr(34)
            oWShell.Run sCmd, 0, True
            Set Folder = Nothing

            ' schedule an additional run of the tool after reboot
            If NOT fRerun Then Rerun

            ' schedule folder for delete on next reboot
            ScheduleDeleteFolder sDelFolder
        End If 'Err <> 0
    Else
        LogOnly "Preview mode. Disallowing delete of folder: " & sDelFolder
    End If
    On Error Goto 0
End Sub 'DeleteFolder

Sub DeleteFolder_WMI (sFolder)
    Dim Folder, Folders
    Dim sWqlFolder
    Dim iRet

    sWqlFolder = Replace(sFolder, "\", "\\")
    Set Folders = oWmiLocal.ExecQuery ("Select * from Win32_Directory where name = '" & sWqlFolder & "'")
    For Each Folder in Folders
        iRet = Folder.Delete
    Next 'Folder
    LogOnly "   Delete (wmi) for folder " & sFolder & " returned: " & iRet
End Sub

'-------------------------------------------------------------------------------
'   Wow64Folder
'
'   Returns the WOW folder structure to handle folder-path operations on
'   64 bit environments
'-------------------------------------------------------------------------------
Function Wow64Folder(sFolder)
    If LCase(Left(sFolder, Len(sWinDir & "\System32"))) = LCase(sWinDir & "\System32") Then 
        Wow64Folder = sWinDir & "\syswow64" & Right(sFolder, Len(sFolder) - Len(sWinDir & "\System32"))
    ElseIf LCase(Left(sFolder, Len(sProgramFiles))) = LCase(sProgramFiles) Then 
        Wow64Folder = sProgramFilesX86 & Right(sFolder, Len(sFolder) - Len(sProgramFiles))
    Else
        Wow64Folder = "?" 'Return invalid string to ensure the folder cannot exist
    End If
End Function 'Wow64Folder

'-------------------------------------------------------------------------------
'   ScheduleDeleteFile
'
'   Adds a file to the list of items to delete on reboot
'-------------------------------------------------------------------------------
Sub ScheduleDeleteFile (sFile)
    If NOT dicDelInUse.Exists(sFile) Then dicDelInUse.Add sFile, sFile Else Exit Sub
    LogOnly "Add file in use for delete on reboot: " & sFile
    fRebootRequired = True
    SetError ERROR_REBOOT_REQUIRED
End Sub 'ScheduleDeleteFile

'-------------------------------------------------------------------------------
'   ScheduleDeleteFolder
'
'   Recursively adds a folder and its contents to the list of 
'   items to delete on reboot
'-------------------------------------------------------------------------------
Sub ScheduleDeleteFolder (sFolder)
    Dim oFolder, fld, file, attr

	Set oFolder = oFso.GetFolder(sFolder)
	' exclude hidden system  folders
    attr = oFolder.Attributes
    If CBool(attr AND 6) Then Exit Sub

	For Each fld In oFolder.SubFolders
		DeleteFolder fld.Path
	Next
	For Each file In oFolder.Files
		DeleteFile file.Path
	Next
	If NOT dicDelInUse.Exists(oFolder.Path) Then dicDelInUse.Add oFolder.Path, "" Else Exit Sub
    LogOnly "Add folder for delete on reboot: " & oFolder.Path
    fRebootRequired = True
    SetError ERROR_REBOOT_REQUIRED
End Sub 'ScheduleDeleteFile


'-------------------------------------------------------------------------------
'   ScheduleDeleteEx
'
'   Schedules the delete of files/folders in use on next reboot by adding
'   affected files/folders to the PendingFileRenameOperations registry entry
'-------------------------------------------------------------------------------
Sub ScheduleDeleteEx ()
    Dim key, hDefKey, sKeyName, sValueName
    Dim i
    Dim arrData

    hDefKey = HKLM
    sKeyName = "SYSTEM\CurrentControlSet\Control\Session Manager"
    sValueName = "PendingFileRenameOperations"
    
    LogH2 "Add " & dicDelInUse.Count & " PendingFileRenameOperations"
    If NOT RegValExists(hDefKey, sKeyName, sValueName) Then
        ReDim arrData(-1)
    Else
        oReg.GetMultiStringValue hDefKey, sKeyName, sValueName, arrData
    End If
    i = UBound(arrData) + 1
    ReDim Preserve arrData(UBound(arrData) + (dicDelInUse.Count * 2))
    For Each key in dicDelInUse.Keys
        LogOnly "   " & key
        arrData(i) = "\??\" & key
        arrData(i + 1) = ""
        i = i + 2
    Next 'key
    oReg.SetMultiStringValue hDefKey, sKeyName, sValueName, arrData
End Sub 'ScheduleDeleteEx

'-------------------------------------------------------------------------------
'   DeleteEmptyFolders
'
'   Deletes an individual folder structure if empty
'-------------------------------------------------------------------------------
Sub DeleteEmptyFolder (sFolder)
    Dim Folder
    
    ' cosmetic' task don't fail on error
    On Error Resume Next
    If oFso.FolderExists(sFolder) Then
        Set Folder = oFso.GetFolder(sFolder)
        If (Folder.Subfolders.Count = 0) AND (Folder.Files.Count = 0) Then 
            Set Folder = Nothing
            SmartDeleteFolder sFolder
        End If
    End If
    CheckError "DeleteEmptyFolder"
    On Error Goto 0
End Sub 'DeleteEmptyFolders

'-------------------------------------------------------------------------------
'   DeleteEmptyFolders
'
'   Delete an empty folder structure
'-------------------------------------------------------------------------------
Sub DeleteEmptyFolders
    Dim Folder
    Dim sFolder
    
    ' cosmetic' task don't fail on error
    On Error Resume Next
    DeleteEmptyFolder sCommonProgramFiles & "\Microsoft Shared\Office15"
    DeleteEmptyFolder sCommonProgramFiles & "\Microsoft Shared\Office16"
    DeleteEmptyFolder sCommonProgramFiles & "\Microsoft Shared\" 
    DeleteEmptyFolder sProgramFiles & "\Microsoft Office\Office15"
    DeleteEmptyFolder sProgramFiles & "\Microsoft Office\Office16"
    
    For Each sFolder in dicDelFolder.Keys
        If oFso.FolderExists(sFolder) Then
            Set Folder = oFso.GetFolder(sFolder)
            If (Folder.Subfolders.Count = 0) AND (Folder.Files.Count = 0) Then 
                Set Folder = Nothing
                SmartDeleteFolder sFolder
            End If
        End If
    Next 'sFolder
    CheckError "DeleteEmptyFolders"
    On Error Goto 0
End Sub 'DeleteEmptyFolders

'-------------------------------------------------------------------------------
'   SmartDeleteFolder
'
'   Wrapper to delete a folder and the empty parent folder structure
'-------------------------------------------------------------------------------
Sub SmartDeleteFolder(sFolder)
    Dim sDelFolder

    If oFso.FolderExists(sFolder) Then
        sDelFolder = sFolder
    ElseIf f64 AND oFso.FolderExists(Wow64Folder(sFolder)) Then
        sDelFolder = Wow64Folder(sFolder)
    Else
        Exit Sub
    End If

    If Not fDetectOnly Then
        LogOnly "Request SmartDelete for folder: " & sDelFolder
        SmartDeleteFolderEx sDelFolder
    Else
        LogOnly "Preview mode. Disallowing SmartDelete request for folder: " & sDelFolder
    End If
End Sub 'SmartDeleteFolder

'-------------------------------------------------------------------------------
'   SmartDeleteFolderEx
'
'   Executes the folder delete operation(s)
'-------------------------------------------------------------------------------
Sub SmartDeleteFolderEx(sFolder)
    Dim Folder
    
    On Error Resume Next
    DeleteFolder sFolder : CheckError "SmartDeleteFolderEx"
    On Error Goto 0
    Set Folder = oFso.GetFolder(oFso.GetParentFolderName(sFolder))
    If (Folder.Subfolders.Count = 0) AND (Folder.Files.Count = 0) Then SmartDeleteFolderEx(Folder.Path)
End Sub 'SmartDeleteFolderEx

'-------------------------------------------------------------------------------
'   RestoreExplorer
'
'   Ensure Windows Explorer is restarted if needed
'-------------------------------------------------------------------------------
Sub RestoreExplorer
    Dim Processes, Result, oAT, DateTime, JobID
    Dim sCmd
    
    'Non critical routine. Don't fail on error
    On Error Resume Next
    wscript.sleep 1000
    Set Processes = oWmiLocal.ExecQuery("Select * From Win32_Process Where Name='explorer.exe'")
    If Processes.Count < 1 Then 
        oWShell.Run "explorer.exe"
        'To handle this in case of System context, schedule and run as interactive task
        oWShell.Run "SCHTASKS /Create /TN OffScrEx /TR explorer /SC ONCE /ST 12:00 /IT", 0, True
        oWShell.Run "SCHTASKS /Run /TN OffScrEx", 0, True
        oWShell.Run "SCHTASKS /Delete /TN OffScrEx /F", 0, False
    End If
    On Error Goto 0
End Sub 'RestoreExploer

'-------------------------------------------------------------------------------
'   MyJoin
'
'   Replacement function to the internal Join function to prevent failures
'   that were seen in some instances
'-------------------------------------------------------------------------------
Function MyJoin(arrToJoin, sSeparator)
    Dim sJoined
    Dim i

    sJoined = ""
    If IsArray(arrToJoin) Then
        For i = 0 To UBound(arrToJoin)
            sJoined = sJoined & arrToJoin(i) & sSeparator
        Next 'i
    End If
    If Len(sJoined) > 1 Then sJoined = Left(sJoined, Len(sJoined) - 1)
    MyJoin = sJoined
End Function

'-------------------------------------------------------------------------------
'   Rerun
'
'   Flag need for reboot and schedule autorun to run the tool again on reboot.
'-------------------------------------------------------------------------------
Sub Rerun ()
    Dim sValue

    ' check if Rerun has already been called
    If fRerun Then Exit Sub

    ' set Rerun flag
    fRerun = True

    ' check if the previous run already initiated the Rerun
    If RegReadValue(HKCU, "SOFTWARE\Microsoft\Office\15.0\CleanC2R", "Rerun", sValue, "REG_DWORD") Then
        ' Rerun has already been tried
        LogH2 "Error: Removal failed"
        SetError ERROR_DCAF_FAILURE
        Exit Sub
    End If

    fRebootRequired = True
    SetError ERROR_REBOOT_REQUIRED
    SetError ERROR_INCOMPLETE

    ' cache the script to the local scrub folder
    oFso.CopyFile WScript.scriptFullName, sScrubDir & "\" & SCRIPTFILE

    oReg.CreateKey HKLM, "SOFTWARE"
    oReg.CreateKey HKLM, "SOFTWARE\Microsoft"
    oReg.CreateKey HKLM, "SOFTWARE\Microsoft\Office"
    oReg.CreateKey HKLM, "SOFTWARE\Microsoft\Office\15.0"
    oReg.CreateKey HKLM, "SOFTWARE\Microsoft\Office\15.0\CleanC2R"
    oReg.SetDWordValue HKLM, "SOFTWARE\Microsoft\Office\15.0\CleanC2R", "Rerun", 1

    fSetRunOnce = True
'    oReg.CreateKey HKCU, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"
'    oReg.SetStringValue HKCU, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce", "CleanC2R", "cscript.exe " & chr(34) & sScrubDir & "\" & SCRIPTFILE & chr(34)
End Sub

'-------------------------------------------------------------------------------
'   SetRunOnce
'
'   Create a RunOnce entry to resume setup after a reboot
'-------------------------------------------------------------------------------
Sub SetRunOnce
    Dim sValue

    oReg.CreateKey HKLM, "SOFTWARE\Microsoft\Windows\CurrentVersion"
    oReg.CreateKey HKLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"
    sValue = "cscript.exe " & chr(34) &  sScrubDir & "\" & SCRIPTFILE & chr(34) & " /NoElevate /Relaunched"
    oReg.SetStringValue HKLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce", "O15CleanUp", sValue

End Sub 'SetRunOnce
