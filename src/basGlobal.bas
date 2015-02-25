Option Compare Database
Option Explicit

' (c) 2000 - 2005 adaept Peter Ennis
' 09/08/2004 1.0.1 - Add code to minimize Access on startup.
' 02/23/2005 1.0.2 - Office 2003 Compatability, GetAccessVersion()
' 02/25/2005 1.0.3 - Use tblAppSetup, gintApp, DLookup for flexibility
' 03/03/2005 1.0.4 - Debug operation with Medical database


' GLOBAL CONSTANTS
Public Const gconTHIS_APP_VERSION As String = "1.0.4"
Public Const gconTHIS_APP_VERSION_DATE = "03/03/2005"
Public Const gconTHIS_APP_NAME = "adaept db loader"
Public gintApp As Integer
Public gstrTheAppWindowName As String
Public gstrTheApp As String
Public gstrTheAppExtension As String
Public gstrTheWorkgroup As String
Public gstrLocalPath As String
Public gstrUpdateAppFile As String
'
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_CLOSE = &H10
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
        ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
'
Global Const SW_HIDE = 0
Global Const SW_SHOWNORMAL = 1
Global Const SW_SHOWMINIMIZED = 2
Global Const SW_SHOWMAXIMIZED = 3
Private Declare Function apiShowWindow Lib "user32" Alias "ShowWindow" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
'

Public Function StartApp() As Boolean

    Dim strTheFile As String
    
    ' Shutdown the app if it is already open
'MsgBox "S1 gintApp = " & gintApp
    gstrTheAppWindowName = DLookup("gconAPP_WINDOW_NAME", "tblAppSetup", _
                            "AppID=" & gintApp)
    gstrLocalPath = DLookup("gconLOCAL_PATH", "tblAppSetup", _
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
'MsgBox "gstrTheAppWindowName = " & gstrTheAppWindowName & vbCrLf & _
'            "gstrTheApp = " & gstrTheApp
    ShutDownApplication (gstrTheAppWindowName)
'MsgBox "S2"
    '
    strTheFile = gstrLocalPath & gstrUpdateAppFile
    
'MsgBox "S3 strTheFile = " & strTheFile
    StartApp = LoadApp(strTheFile)
'Exit Function
'MsgBox "S4"

    DoCmd.Quit

End Function

Private Function LoadApp(strAbsAppName As String) As Boolean
' What:         Load the selected application
' Author:       Peter F. Ennis
' Created:      8/2004
' Passed in:    Absolute application file name as a string
' Returns:      True if successful
' Last Mod:

On Error GoTo Err_LoadApp        ' Set up error handler.

'MsgBox "L1"
    If FileExists(strAbsAppName) Then
'MsgBox "L2"
        ' Rename the old app file
        Name Mid(strAbsAppName, 1, Len(strAbsAppName) - 3) & gstrTheAppExtension _
                As Mid(strAbsAppName, 1, Len(strAbsAppName) - 3) & "OLD"
        ' Rename the update app file
'MsgBox "L3"
        Name strAbsAppName As Mid(strAbsAppName, 1, _
                Len(strAbsAppName) - 3) & gstrTheAppExtension
    End If
'MsgBox "L4 gstrTheAppWindowName = " & gstrTheAppWindowName & vbCrLf & _
'            "gstrTheApp = " & gstrTheApp
'Exit Function

    Dim i As Integer
    Do
'MsgBox "L5"
        OpenSecured gstrTheApp, gstrTheWorkgroup, "IntakeUser", "dscc"
        
        i = MsgBox("L6", vbYesNo, "Test Break")
        If i = vbYes Then
            Exit Function
        Else
        End If
        
        DoEvents
'MsgBox "L7"
    Loop Until WindowIsOpen(gstrTheAppWindowName)
'MsgBox "L8"
    'MsgBox WindowIsOpen("Davis Street Family Resource Center")
    'MaximizeTheWindow WindowIsOpen("Davis Street Family Resource Center"), "Davis Street Family Resource Center"
    
    LoadApp = True
    
Exit_LoadApp:
    Exit Function

Err_LoadApp:
    'MsgBox Err & " " & Err.Description, vbCritical, "LoadApp: " & gconTHIS_APP_NAME
    Select Case Err
        Case 58
            ' OLD app file exists
            Kill Mid(strAbsAppName, 1, Len(strAbsAppName) - 3) & "OLD"
            Resume
        Case 75
            ' Path/File access error: If app is open it takes time to be
            ' shut down so try again
            Delay 1
            Resume
        Case Else
            MsgBox Err & " " & Err.Description, vbCritical, gconTHIS_APP_NAME & ": LoadApp"
    End Select
    LoadApp = False
    Resume Exit_LoadApp

End Function

Private Function FileExists(strAbsFileName As String) As Boolean
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

Private Function ShutDownApplication(ByVal strApplicationName As String) As Boolean
'Ref: http://www.a1vbcode.com/app.asp?ID=479

    Dim hwnd As Long
    Dim Result As Long
    hwnd = FindWindow(vbNullString, strApplicationName)
    'MsgBox "hWnd = " & hWnd & vbCrLf & _
    '            "strApplicationName = " & strApplicationName
    If hwnd <> 0 Then
        Result = PostMessage(hwnd, WM_CLOSE, 0&, 0&)
        ShutDownApplication = True
        'MsgBox "The application window was found and shutdown."
    Else
        MsgBox "The application window " & _
            strApplicationName & " was not found."
    End If

End Function

Private Function WindowIsOpen(ByVal strWindowTitle As String) As Long

    Dim hwnd As Long
    Dim Result As Long
    hwnd = FindWindow(vbNullString, strWindowTitle)
    Debug.Print "hwnd = " & hwnd
    If hwnd <> 0 Then
        WindowIsOpen = hwnd
    Else
        WindowIsOpen = 0
    End If

End Function

Private Sub MaximizeTheWindow(hwnd As Long, ByVal strWindowTitle As String)
'Ref: http://www.digital-inn.de/archive/index.php/t-15364.html

    Dim lng As Long
    lng = SendMessage(hwnd, &H112, &HF030&, 0&)
    
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
'MsgBox "OSec1"
    Set objAccess = GetObject(, "Access.Application")
'MsgBox "OSec2 Err = " & Err
    If Err = 0 Then 'an instance of Access is open
'MsgBox "OSec3 Err = " & Err
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

'MsgBox "OSec4 Err = " & Err
    If GetAccessVersion = "9.0" Then        ' Access 2000
        cmd = """C:\Program Files\Microsoft Office\Office\MSAccess.exe""" & " "
    ElseIf GetAccessVersion = "11.0" Then   ' Access 2003
        cmd = """C:\Program Files\Microsoft Office\Office11\MSAccess.exe""" & " "
    End If

    cmd = cmd & """" & strTheApp & """" & " " & _
                 "/wrkgrp" & " " & _
                 """" & strTheWorkgroup & """"
    'MsgBox cmd
'Exit Sub
        '
'MsgBox "OSec5 Err = " & Err
        cmd = cmd & " /nostartup /user " & varUser
'MsgBox "OSec6 Err = " & Err & " cmd = " & cmd
        If Not IsMissing(varPw) Then cmd = cmd & " /pwd " & varPw
'MsgBox "OSec7"
        Shell pathname:=cmd, windowstyle:=6
        Dim bln As Boolean
        bln = fIsAppRunning("access")
'MsgBox "OSec8"
        Do 'Wait for shelled process to finish.
'MsgBox "OSec9"
            Err = 0
            Set objAccess = GetObject(, "Access.Application")
        Loop While Err <> 0
'MsgBox "OSec10"
    End If

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

'Private Function aeGetCmdString(intAPP As Integer) As String
'
'    Dim conAPP_SERVER_PATH As String        ' "\\Dscc-w2k-1\Intake\"
'    Dim conAPP_LOCAL_PATH As String         ' "C:\DSFRC\Intake\"
'    Dim conAPP_UPDATE_INFO_FILE As String   ' "DSFRC Update Info.txt"
'    Dim conAPP_UPDATE_APP_FILE As String    ' "Davis Street Intake PRODUCTION SQL 2000 Front End A2K.upd"
'
'End Function

Public Function aeGetTheAppID() As Integer
' Ref: http://msdn.microsoft.com/library/default.asp?url=/library/en-us/dnacc2k/html/acglobaloptions.asp

    aeGetTheAppID = Nz(DLookup("AppID", "tblAppSetup", _
                    "gconAPP_CMD_NAME='" & Command & "'"), 0)
    'MsgBox "aeGetTheAppID = " & aeGetTheAppID

'    If Command = "DSFRC Intake" Then
'        aeGetTheAppID = 1
'        Exit Function
'    End If
'
'    If Command = "DSFRC Medical" Then
'        aeGetTheAppID = 2
'        Exit Function
'    End If

    If aeGetTheAppID = 0 Then
        MsgBox "Invalid Access Command Line Parameter!" & vbCrLf & vbCrLf & _
                Command, vbCritical, gconTHIS_APP_NAME
        DoCmd.Quit acQuitSaveNone
    End If

End Function