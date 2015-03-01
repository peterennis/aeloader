Option Compare Database
Option Explicit

' (c) 2000 - 2005 adaept Peter Ennis
' 09/08/2004 1.0.1 - Add code to minimize Access on startup.
' 02/23/2005 1.0.2 - Office 2003 Compatability, GetAccessVersion()
' 02/25/2005 1.0.3 - Use tblAppSetup, gintApp, DLookup for flexibility
' 03/03/2005 1.0.4 - Debug operation with Medical database
' 03/21/2005 1.0.5 - Add SLCC with Logon and Password fields in startup table.
' 07/22/2005 1.0.6 - DoCmd.Restore added when application closes.
' 08/18/2005 1.0.7 - Add DSFRC Volunteers.
' 08/24/2005 1.0.8 - Add DSFRC Finance.
' 08/31/2005 1.0.9 - Add NOPWD group and NopwdUser as pass through for Noho.
' 09/15/2005 1.1.0 - Pass through test succeeds for NoHo.
' 09/16/2005 1.1.1 - Fix bug where pass through messed up original version control.
' 06/09/2006 1.1.2 - Add aeLoaderMoveSizeClass to center and reduce access db window
'                    to its smallest size on any screen.
' 05/10/2007 1.1.3 - Allow update with user:PCname and kill .OLD mda file before rename.
' 05/18/2007 1.1.4 - Updates to class modules from adaeptdblib.mda
'                    Debugging to trap file permissions error when user is not in admin group.
'                    Need to allow the application to delete and rename files.
'

' GLOBAL CONSTANTS
Public Const gconTHIS_APP_VERSION As String = "1.1.4"
Public Const gconTHIS_APP_VERSION_DATE = "05/18/2007"
Public Const gconTHIS_APP_NAME = "adaept db loader"
Public gblnAbortUpdate As Boolean
Public gblnSPAWN_DEBUG As Boolean
Public gintApp As Integer
Public gstrTheAppWindowName As String
Public gstrTheApp As String
'
Public gstrAppName As String
Public gstrAppCurrentVer As String
Public gstrDbLibVersion As String
Public gstrAppNewVersion As String
Public gstrLocalLibPath As String
Public gstrDbLibName As String
Public gstrUpdateMdb As String
'
Public gstrTheNewApp As String
Public gstrTheAppNamePart As String
Public gstrTheAppVersionPart As String
Public gstrTheAppSeparatorChar As String
Public gstrPassThrough As String
Public gstrTheAppExtension As String
Public gstrAppCmdName As String
Public gstrTheWorkgroup As String
Public gstrUpdateAppFile As String
Public gstrLoaderUpdateAppFile As String
Public gstrLogonMdb As String
Public gstrPasswordMdb As String
'
' GLOBAL CONSTANTS FOR PASS_THROUGH UPDATE
Public gfUpdateDebug As Boolean
Public gstrTheCurrentUser As String
Public gstrComputerName As String
Public gstrLocalPath As String
Public gstrServerPath As String
Public gstrUpdateInfoFile As String
Public gstrDebugFile As String
Public gstrNetUserLogin As String
Public gstrMdbUserLogin As String
Public gstrSqlUserLogin As String
Public gstrAppCurrentFileVer As String
Public gstrUpdateText As String
Public gstrAppNewFileVersion As String
'
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_CLOSE = &H10
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, _
        ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
'
Global Const SW_HIDE = 0
Global Const SW_SHOWNORMAL = 1
Global Const SW_SHOWMINIMIZED = 2
Global Const SW_SHOWMAXIMIZED = 3
Private Declare Function apiShowWindow Lib "user32" Alias "ShowWindow" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
'

Public Function StartApp() As Boolean

    Dim strTheFile As String

On Error GoTo Err_StartApp

    gstrPassThrough = Nz(DLookup("gconPASS_THROUGH", "tblAppSetup", _
                            "AppID=" & gintApp))
    'MsgBox "gstrPassThrough = " & gstrPassThrough
    If gstrPassThrough = "PassThrough" Then
        ' Call aeLoaderUpdateSetupClass
        Dim blnUpdate As Boolean
        Dim cls1 As aeLoaderUpdateSetupClass
        Set cls1 = New aeLoaderUpdateSetupClass
        ' Setup parameters
        cls1.aeUpdateDebug = True
        blnUpdate = cls1.aeUpdateSetup(gconTHIS_APP_NAME, gconTHIS_APP_VERSION, aeWindowsNetworkLogin)
    
        ' DSFRC with Network Login
        Dim cls2 As aeLoaderUpdateTxtClass
        Set cls2 = New aeLoaderUpdateTxtClass
        cls2.aeUpdateDebug = True
        blnUpdate = cls2.blnTheAppLoaderUpdateStatus()
        Debug.Print "cls2.blnTheAppLoaderUpdateStatus = " & blnUpdate
        '
        ' Shutdown the app if it is already open
        'MsgBox "gstrTheAppWindowName = " & gstrTheAppWindowName, vbInformation, gconTHIS_APP_NAME & ": StartApp"
        ShutDownApplication (gstrTheAppWindowName)
        '
        StartApp = aeLoaderPassThroughApp(gstrLocalPath, gstrLoaderUpdateAppFile)
        DoCmd.Quit
        Exit Function
    End If
        
    ' Minimize the Access window
    ShowWindow Application.hWndAccessApp, 2
    
    ' UPDATE PROCESS USING adaeptdblib.mda LINKED
    ' IN MDB APPLICATION
    '
    ' Shutdown the app if it is already open
    gstrTheAppWindowName = DLookup("gconAPP_WINDOW_NAME", "tblAppSetup", _
                            "AppID=" & gintApp)
    gstrLocalPath = DLookup("gconLOCAL_PATH", "tblAppSetup", _
                            "AppID=" & gintApp)
    gstrUpdateAppFile = DLookup("gconUPDATE_APP_FILE", "tblAppSetup", _
                            "AppID=" & gintApp)
    gstrTheAppExtension = DLookup("gconAPP_EXTENSION", "tblAppSetup", _
                            "AppID=" & gintApp)
    gstrTheWorkgroup = DLookup("gconSERVER_PATH", "tblAppSetup", _
                            "AppID=" & gintApp) & _
                            DLookup("gconTHE_WORKGROUP", "tblAppSetup", _
                            "AppID=" & gintApp)
    gstrTheApp = gstrLocalPath & DLookup("gconAPP_NAME", "tblAppSetup", _
                            "AppID=" & gintApp) & _
                            "." & gstrTheAppExtension
    gstrLogonMdb = DLookup("gconLOGON_MDB", "tblAppSetup", _
                            "AppID=" & gintApp)
    gstrPasswordMdb = DLookup("gconPASSWORD_MDB", "tblAppSetup", _
                            "AppID=" & gintApp)
    'MsgBox "gstrTheAppWindowName = " & gstrTheAppWindowName & vbCrLf & _
            "gstrLocalPath = " & gstrLocalPath & vbCrLf & _
            "gstrUpdateAppFile = " & gstrUpdateAppFile & vbCrLf & _
            "gstrTheAppExtension = " & gstrTheAppExtension & vbCrLf & _
            "gstrTheWorkgroup = " & gstrTheWorkgroup & vbCrLf & _
            "gstrTheApp = " & gstrTheApp & vbCrLf & _
            "gstrLogonMdb = " & gstrLogonMdb & vbCrLf & _
            "gstrPasswordMdb = " & gstrPasswordMdb
'MsgBox "1"
    ShutDownApplication (gstrTheAppWindowName)
'MsgBox "2"
    '
    ' Update to new library
    If FileExists("C:\DSFRC\Intake\adaeptdblib.mda.upd") Then
'MsgBox "3"
        Kill "C:\DSFRC\Intake\adaeptdblib.mda.OLD"
        Name "C:\DSFRC\Intake\adaeptdblib.mda" _
                As "C:\DSFRC\Intake\adaeptdblib.mda." & "OLD"
        ' Rename the update app file
'MsgBox "4"
        Name "C:\DSFRC\Intake\adaeptdblib.mda.upd" _
                As "C:\DSFRC\Intake\adaeptdblib.mda"
    End If
    
'MsgBox "5"
    strTheFile = gstrLocalPath & gstrUpdateAppFile
    'MsgBox "StartApp: strTheFile = " & strTheFile
'MsgBox "6"
    StartApp = aeLoaderApp(strTheFile)
'MsgBox "7"

    DoCmd.Restore
    DoCmd.Quit

Exit_StartApp:
    Exit Function

Err_StartApp:
    Select Case Err
'          Case 58
'            ' OLD app file exists
'            Kill Mid(strAbsAppName, 1, Len(strAbsAppName) - 3) & "OLD"
'            Resume
          Case 75
          ' Path/File access error: If app is open it takes time to be
            ' shut down so try again
            Delay 1
            Resume
        Case Else
            MsgBox "Erl:" & Erl & " Error# " & Err & " " & Err.Description, vbCritical, "aeLoaderPassThroughApp: " & gconTHIS_APP_NAME
    End Select
    Resume Exit_StartApp

End Function

Public Function aeLoaderPassThroughApp(strPath As String, strFileName As String) As Boolean
' What:         Load the selected pass through application
' Author:       Peter F. Ennis
' Created:      9/13/2003
' Passed in:    Absolute application file name as a string
' Returns:      True if successful
' Last Mod:
    
On Error GoTo Err_aeLoaderPassThroughApp        ' Set up error handler.

1    If FileExists(strPath & strFileName) Then
2        'MsgBox strPath & strFileName & " FOUND." & vbCrLf & _
            "WRITE CODE TO KILL OLD APPS", vbInformation, "aeLoaderPassThroughApp: " & gconTHIS_APP_NAME
         Debug.Print ">aeLoaderPassThroughApp: strPath = " & strPath
         Debug.Print ">aeLoaderPassThroughApp: strFileName = " & strFileName
3        KillOldApps strPath, strFileName
4    End If

5    Do
6        OpenNotSecured strPath & strFileName  ', gstrTheWorkgroup
        
7        If gblnSPAWN_DEBUG Then
8            Dim i As Integer
9            i = MsgBox("L6 gblnSPAWN_DEBUG", vbYesNo, "Test Break")
10            If i = vbYes Then
11                Exit Function
12            Else
13            End If
14        End If
        
15        DoEvents
16    Loop Until WindowIsOpen(gstrTheAppWindowName)

17    aeLoaderPassThroughApp = True
    
Exit_aeLoaderPassThroughApp:
18    Exit Function

Err_aeLoaderPassThroughApp:
    'MsgBox Err & " " & Err.Description, vbCritical, "aeLoaderPassThroughApp: " & gconTHIS_APP_NAME
19    Select Case Err
          Case 75
20          ' Path/File access error: If app is open it takes time to be
            ' shut down so try again
21            Delay 1
22            Resume
23        Case Else
24            MsgBox "Erl:" & Erl & " Error# " & Err & " " & Err.Description & vbCrLf & _
                    "strPath = " & strPath & _
                    "strFileName = " & strFileName, vbCritical, "aeLoaderPassThroughApp: " & gconTHIS_APP_NAME
25    End Select
26    aeLoaderPassThroughApp = False
27    Resume Exit_aeLoaderPassThroughApp

End Function

Public Sub KillOldApps(strPath As String, strFileName As String)

On Error GoTo Err_KillOldApps

    Dim strFName As String
    Dim strFilePattern As String
    
1    Debug.Print "strPath = " & strPath
2    Debug.Print "strFileName = " & strFileName
3    strFilePattern = Left(strFileName, InStr(strFileName, gstrTheAppSeparatorChar))
4    Debug.Print "strFilePattern = " & strFilePattern
    
    ' Display the names in strPath that represent the application to be started
5    strFName = Dir(strPath & strFilePattern & "*")    ' Retrieve the first entry.
6    Do While strFName <> ""    ' Start the loop.
7         If strFName <> strFileName Then
8             Debug.Print "Found: " & strFName
9             Kill strPath & strFName
10        Else
11            Debug.Print "APP TO LOAD: " & strFName
12        End If
13        strFName = Dir    ' Get next entry.
14    Loop
15    'Stop
      ' Make copy of app bmp startup file
      If FileExists(strPath & gstrAppCmdName & ".bmp") Then
            'MsgBox "Creating App bmp File"
            FileCopy strPath & gstrAppCmdName & ".bmp", strPath & Mid(strFileName, 1, Len(strFileName) - 4) & ".bmp"
      End If

Exit_KillOldApps:
Exit Sub

Err_KillOldApps:
    MsgBox "Erl:" & Erl & " Error# " & Err & " " & Err.Description, vbCritical, "KillOldApps: " & gconTHIS_APP_NAME
    Resume Next

End Sub

Private Function aeLoaderApp(strAbsAppName As String) As Boolean
' What:         Load the selected application
' Author:       Peter F. Ennis
' Created:      8/2004
' Passed in:    Absolute application file name as a string
' Returns:      True if successful
' Last Mod:     09/06/2005 Use line numbers and Erl to help debugging

On Error GoTo Err_aeLoaderApp        ' Set up error handler.
     
1    If FileExists(strAbsAppName) Then
        ' Rename the old app file
2        Name Mid(strAbsAppName, 1, Len(strAbsAppName) - 3) & gstrTheAppExtension _
                As Mid(strAbsAppName, 1, Len(strAbsAppName) - 3) & "OLD"
        ' Rename the update app file
3        Name strAbsAppName As Mid(strAbsAppName, 1, _
                Len(strAbsAppName) - 3) & gstrTheAppExtension
4    End If

5    Do
6        OpenSecured gstrTheApp, gstrTheWorkgroup, gstrLogonMdb, gstrPasswordMdb
        
7        If gblnSPAWN_DEBUG Then
8            Dim i As Integer
9            i = MsgBox("L6", vbYesNo, "Test Break")
10            If i = vbYes Then
11                Exit Function
12            Else
13            End If
14        End If
        
15        DoEvents
16    Loop Until WindowIsOpen(gstrTheAppWindowName)
    'MsgBox WindowIsOpen("Davis Street Family Resource Center")
    'MaximizeTheWindow WindowIsOpen("Davis Street Family Resource Center"), "Davis Street Family Resource Center"

17    aeLoaderApp = True
    
Exit_aeLoaderApp:
18    Exit Function

Err_aeLoaderApp:
    'MsgBox Err & " " & Err.Description, vbCritical, "aeLoaderApp: " & gconTHIS_APP_NAME
19    Select Case Err
        Case 58
            ' OLD app file exists
20            Kill Mid(strAbsAppName, 1, Len(strAbsAppName) - 3) & "OLD"
21            Resume
22        Case 75
            ' Path/File access error: If app is open it takes time to be
            ' shut down so try again
23            Delay 1
24            Resume
25        Case Else
26            MsgBox "Erl:" & Erl & " Error# " & Err & " " & Err.Description, vbCritical, gconTHIS_APP_NAME & ": aeLoaderApp"
27    End Select
28    aeLoaderApp = False
29    Resume Exit_aeLoaderApp

End Function

Public Function FileExists(strAbsFileName As String) As Boolean
' What:         Test for existence of a file.
' Author:       Peter F. Ennis
' Created:      11/1998
' Passed in:    Absolute file name as a string
' Returns:      True
' Last Mod:     07/30/99
'               08/27/2004 use gconTHIS_APP_NAME in messages

    Dim strSubName As String
    strSubName = "FileExists"
    
    Dim strMessage As String
   
On Error GoTo Err_FileExists        ' Set up error handler.

    FileExists = (Dir(strAbsFileName) <> "")

Exit_FileExists:
    Exit Function

Err_FileExists:
    Select Case Err
        Case 53, 62
            'ADD FURTHER DESCRIPTION AS DESIRED
            MsgBox "FileExists Error # " & Err.Number & ": " & Err.Description
        Case 68
            strMessage = "FileExists Error # " & Err.Number & ": " & Err.Description & vbCrLf  'Device unavailable
            strMessage = strMessage & "Path: " & strAbsFileName & vbCrLf
            strMessage = strMessage & "Possible cause: Network connection problem or server may be down!"
            MsgBox strMessage, vbInformation, gconTHIS_APP_NAME & ": FileExists"
        Case 76
            strMessage = "FileExists Error # " & Err.Number & ": " & Err.Description & vbCrLf  'Path Not Found
            strMessage = strMessage & "Path: " & strAbsFileName
            MsgBox strMessage, vbInformation, gconTHIS_APP_NAME & ": FileExists"
        Case Else
            strMessage = "Unexpected FileExists Error " & Err.Number & ": " & Err.Description
            MsgBox strMessage, vbCritical, gconTHIS_APP_NAME & ": FileExists"
        End Select
    Resume Exit_FileExists

End Function

Public Function ShutDownApplication(ByVal strApplicationName As String) As Boolean
'Ref: http://www.a1vbcode.com/app.asp?ID=479

    Dim hWnd As Long
    Dim Result As Long
    hWnd = FindWindow(vbNullString, strApplicationName)
    'MsgBox "hWnd = " & hWnd & vbCrLf & _
    '            "strApplicationName = " & strApplicationName
    If hWnd <> 0 Then
        Result = PostMessage(hWnd, WM_CLOSE, 0&, 0&)
        'MsgBox "The application window was found for shutdown."
        ShutDownApplication = True
        ' If the app is NoHo it will not shutdown but give a Quit message
        ' so close the loader
        If gstrTheAppWindowName = "NoHo CARE" Then
            DoCmd.Quit
        End If
    Else
        'MsgBox "The application window " & _
            strApplicationName & " was not found.", vbInformation, _
            gconTHIS_APP_NAME & ": ShutDownApplication"
        'DoCmd.Quit
    End If

End Function

Private Function WindowIsOpen(ByVal strWindowTitle As String) As Long

    Dim hWnd As Long
    Dim Result As Long
    hWnd = FindWindow(vbNullString, strWindowTitle)
    Debug.Print "hwnd = " & hWnd
    If hWnd <> 0 Then
        WindowIsOpen = hWnd
    Else
        WindowIsOpen = 0
    End If

End Function

Private Sub MaximizeTheWindow(hWnd As Long, ByVal strWindowTitle As String)
'Ref: http://www.digital-inn.de/archive/index.php/t-15364.html

    Dim lng As Long
    lng = SendMessage(hWnd, &H112, &HF030&, 0&)
    
End Sub


'----------------------------------------------------------------------
'Using a Secured Workgroup
'If the Microsoft Access application you want to control uses a
'secured workgroup (System.mdw), you may want to bypass the logon box,
'which asks for a user name and password. The following sample code uses
'the Shell() function to start Microsoft Access and pass a user name and
'password to the application:
'----------------------------------------------------------------------
'DECLARATIONS
'----------------------------------------------------------------------
'This procedure sets a module-level variable, objAccess, to refer to
'an instance of Microsoft Access. The code first tries to use GetObject
'to refer to an instance that might already be open. If an instance is
'not already open, the Shell() function opens a new instance and
'specifies the user and password, based on the arguments passed to the
'procedure.
'
'Calling example: OpenSecured varUser:="Admin", varPw:=""
'Ref: http://support.microsoft.com/default.aspx?scid=kb;en-us;210111&Product=acc2000
'----------------------------------------------------------------------
Private Sub OpenSecured(strTheApp As String, _
                            strTheWorkgroup As String, _
                            Optional varUser As Variant, _
                            Optional varPw As Variant)
    
    Dim objAccess As Object
    Dim cmd As String
    
    On Error Resume Next
    Set objAccess = GetObject(, "Access.Application")
    If Err = 0 Then 'an instance of Access is open
        If IsMissing(varUser) Then varUser = "Admin"
        
' ******** EXAMPLE ********
'        cmd = """C:\Program Files\Microsoft Office\Office\MSAccess.exe""" & " " & _
'                 """C:\DSFRC\Intake\Davis Street Intake PRODUCTION SQL 2000 Front End A2K.mdb""" & " " & _
'                 "/wrkgrp" & " " & _
'                 "\\Dscc-w2k-1\Intake\DSCC.MDW" & " " & _
'                 "/cmd " & _
'                 """NOSPLASHFORM"""
'        MsgBox cmd
' **************************

    If GetAccessVersion = "9.0" Then        ' Access 2000
        cmd = """C:\Program Files\Microsoft Office\Office\MSAccess.exe""" & " "
    ElseIf GetAccessVersion = "11.0" Then   ' Access 2003
        cmd = """C:\Program Files\Microsoft Office\Office11\MSAccess.exe""" & " "
    End If

    cmd = cmd & """" & strTheApp & """" & " " & _
                 "/wrkgrp" & " " & _
                 """" & strTheWorkgroup & """"
    'MsgBox cmd
        '
        cmd = cmd & " /nostartup /user " & varUser
        If Not IsMissing(varPw) Then cmd = cmd & " /pwd " & varPw
        Shell pathname:=cmd, windowstyle:=6
        Dim bln As Boolean
        bln = fIsAppRunning("access")
        Do 'Wait for shelled process to finish.
            Err = 0
            Set objAccess = GetObject(, "Access.Application")
        Loop While Err <> 0
    End If

End Sub

Private Sub OpenNotSecured(strTheApp As String)
    
    Dim objAccess As Object
    Dim cmd As String
    
    On Error Resume Next
    Set objAccess = GetObject(, "Access.Application")
            
    If GetAccessVersion = "9.0" Then        ' Access 2000
        cmd = """C:\Program Files\Microsoft Office\Office\MSAccess.exe""" & " "
    ElseIf GetAccessVersion = "11.0" Then   ' Access 2003
        cmd = """C:\Program Files\Microsoft Office\Office11\MSAccess.exe""" & " "
    End If

    cmd = cmd & """" & strTheApp & """"
    'MsgBox cmd, vbInformation, "OpenNotSecured"
    '
    Shell pathname:=cmd, windowstyle:=vbMaximizedFocus
    Dim bln As Boolean
    bln = fIsAppRunning("access")
    Do 'Wait for shelled process to finish.
        Err = 0
        Set objAccess = GetObject(, "Access.Application")
    Loop While Err <> 0

End Sub

Sub Delay(pdblSeconds As Double)
' Delay for x seconds
' This sub uses very little CPU resouces
' Ref: http://www.experts-exchange.com/Programming/Programming_Languages/Visual_Basic/Q_20843293.html
    
    Const OneSecond As Double = 1# / (1440# * 60#)

    Dim dblWaitUntil As Date
    dblWaitUntil = Now + OneSecond * pdblSeconds
    Do Until Now > dblWaitUntil
        Sleep 100
        DoEvents ' Allow windows message to be processed
    Loop

End Sub

Function DoAccessWindow(nCmdShow As Long)
' Ref: http://members.shaw.ca/glenk/access97.html
' Ref: http://www.mvps.org/access/api/api0019.htm
'
' http://support.microsoft.com/?kbid=210090
' Microsoft Knowledge Base Article - 210090
' ACC2000: How to Use Visual Basic for Applications to Minimize, Maximize, and Restore Access

'Usage Examples
'Maximize window:
'       ?DoAccessWindow(SW_SHOWMAXIMIZED)
'Minimize window:
'       ?DoAccessWindow(SW_SHOWMINIMIZED)
'Hide window:
'       ?DoAccessWindow(SW_HIDE)
'Normal window:
'       ?DoAccessWindow(SW_SHOWNORMAL)
'
Dim loX  As Long
Dim loform As Form
    On Error Resume Next
    Set loform = Screen.ActiveForm
    If Err <> 0 Then 'no Activeform
      If nCmdShow = SW_HIDE Then
        MsgBox "Cannot hide Access unless a form is on screen"
      Else
        loX = apiShowWindow(hWndAccessApp, nCmdShow)
        Err.Clear
      End If
    Else
        If nCmdShow = SW_SHOWMINIMIZED And loform.Modal = True Then
            MsgBox "Cannot minimize Access with " & (loform.Caption + " ") & "form on screen"
        ElseIf nCmdShow = SW_HIDE And loform.PopUp <> True Then
            MsgBox "Cannot hide Access with " & (loform.Caption + " ") & "form on screen"
        Else
            loX = apiShowWindow(hWndAccessApp, nCmdShow)
        End If
    End If
    DoAccessWindow = (loX <> 0)

End Function

Function GetAccessVersion() As String
' Ref: http://www.blueclaw-db.com/get_access_version_number.htm
' To determine the version of Microsoft Access used to open this application.
' 8.0 = Access 97
' 9.0 = Access 2000
' 10.0 = Access 2002(XP)
' 11.0= Access 2003

    GetAccessVersion = SysCmd(acSysCmdAccessVer)
    
End Function

Public Function aeGetTheAppID() As Integer
' Ref: http://msdn.microsoft.com/library/default.asp?url=/library/en-us/dnacc2k/html/acglobaloptions.asp

    Dim intAppID As Integer

    'MsgBox "Command = " & Command
    gstrAppCmdName = Command
    intAppID = DLookup("[AppID]", "[tblAppSetup]", _
                        "[gconAPP_CMD_NAME] = '" & gstrAppCmdName & "'")
    'MsgBox "intAppID = " & intAppID

    If intAppID = 0 Then
        MsgBox "Invalid Access Command Line Parameter!" & vbCrLf & vbCrLf & _
                Command, vbCritical, gconTHIS_APP_NAME
        DoCmd.Restore
        DoCmd.Quit acQuitSaveNone
        Exit Function
    End If

    aeGetTheAppID = intAppID
    
'    If Command = "DSFRC Intake" Then
'        aeGetTheAppID = 1
'        Exit Function
'    End If
'
'    If Command = "DSFRC Medical" Then
'        aeGetTheAppID = 2
'        Exit Function
'    End If

End Function

Public Function Comment(strComment As String) As Boolean
' What:         THIS FUNCTION RETURNS TRUE IF A STRING IS A COMMENT i.e. IF
'               THE FIRST CHARACTER IN THE FIRST LINE IS ' OR ;
' Author:       Peter F. Ennis      Created: 11/98       By: Peter F. Ennis
' Passed in:    Comment as a string
' Returns:      True
' Last Mod:     7/30/99

On Error GoTo Err_Comment

    Comment = False
    If ((Mid$(strComment, 1, 1) = "'") Or (Mid$(strComment, 1, 1) = ";")) Then
        Comment = True
    End If

Exit_Comment:
    Exit Function

Err_Comment:
    MsgBox "Comment Error " & Err & ": " & Error$, vbCritical, "aedb"
    Resume Exit_Comment

End Function

Public Function MoveToCenter()

    Dim bln As Boolean
    Dim cls As aeLoaderMoveSizeClass
    Set cls = New aeLoaderMoveSizeClass
    ' Setup parameters
    cls.aeMoveSizeCenter = True

End Function