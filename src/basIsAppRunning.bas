Option Compare Database
Option Explicit

'*****************************************
' Modified by Peter F. Ennis to check for
' multiple copies of an application opening
' so as to prevent uncontrolled spawning.

Private malngAccessHandles() As Long
Private mlngH As Long

' Ref: http://www.mvps.org/access/api/api0007.htm

'***************** Code Start ***************
'This code was originally written by Dev Ashish.
'It is not to be altered or distributed,
'except as part of an application.
'You are free to use it in any application,
'provided the copyright notice is left unchanged.
'
'Code Courtesy of Dev Ashish
'
Private Const SW_HIDE = 0
Private Const SW_SHOWNORMAL = 1
Private Const SW_NORMAL = 1
Private Const SW_SHOWMINIMIZED = 2
Private Const SW_SHOWMAXIMIZED = 3
Private Const SW_MAXIMIZE = 3
Private Const SW_SHOWNOACTIVATE = 4
Private Const SW_SHOW = 5
Private Const SW_MINIMIZE = 6
Private Const SW_SHOWMINNOACTIVE = 7
Private Const SW_SHOWNA = 8
Private Const SW_RESTORE = 9
Private Const SW_SHOWDEFAULT = 10
Private Const SW_MAX = 10

Private Declare PtrSafe Function apiFindWindow Lib "user32" Alias "FindWindowA" (ByVal strClass As String, ByVal lpWindow As String) As Long
Private Declare PtrSafe Function apiSendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, lParam As Long) As Long
Private Declare PtrSafe Function apiSetForegroundWindow Lib "user32" Alias "SetForegroundWindow" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function apiShowWindow Lib "user32" Alias "ShowWindow" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Function apiIsIconic Lib "user32" Alias "IsIconic" (ByVal hWnd As Long) As Long
'

Public Function IsRunning(ByVal strAppName As String, _
        Optional fActivate As Boolean) As Boolean

    On Local Error GoTo PROC_ERR

    Debug.Print "IsRunning"
    Dim strClassName As String
    Dim lngX As Long
    Dim lngTmp As Long

    Const WM_USER = 1024

    IsRunning = False

    Select Case LCase$(strAppName)
        Case "access":      strClassName = "OMain"
        Case "excel":       strClassName = "XLMain"
        Case "notepad":     strClassName = "NOTEPAD"
        Case "outlook":     strClassName = "rctrl_renwnd32"
        Case "paintbrush":  strClassName = "pbParent"
        Case "powerpoint95": strClassName = "PP7FrameClass"
        Case "powerpoint97": strClassName = "PP97FrameClass"
        Case "word":        strClassName = "OpusApp"
        Case "wordpad":     strClassName = "WordPadClass"
        Case Else:          strClassName = vbNullString
    End Select

    Debug.Print , "strClassName = " & strClassName
    If strClassName = "" Then
        mlngH = apiFindWindow(vbNullString, strAppName)
    Else
        mlngH = apiFindWindow(strClassName, vbNullString)
    End If

    If strClassName = "OMain" Then
        Debug.Print , "mlngH = " & mlngH
        malngAccessHandles(GetBounds()) = mlngH
    End If

    If mlngH <> 0 Then
''        apiSendMessage mlngH, WM_USER + 18, 0, 0
''        lngX = apiIsIconic(mlngH)
'        If lngX <> 0 Then
            '#PFE# This line causes aeLoader to maximize
            'lngTmp = apiShowWindow(mlngH, SW_SHOWNORMAL)
'        End If
'        If fActivate Then
''            lngTmp = apiSetForegroundWindow(mlngH)
'        End If
        IsRunning = True
    End If

PROC_EXIT:
    Exit Function

PROC_ERR:
    IsRunning = False
    Resume PROC_EXIT

End Function

Public Function GetBounds() As Integer
' What: Create an array on demand that stores the handles of the open application windows
'  Who: (c) adaept 2005

    On Error GoTo PROC_ERR

    Dim i As Integer

    i = UBound(malngAccessHandles)
    Debug.Print "GetBounds"
    Debug.Print , "i = " & i

    ReDim Preserve malngAccessHandles(i + 1)
    GetBounds = i + 1

    If GetBounds = 3 Then
        'MsgBox "3 Apps Opened! - Halt"
        Debug.Print , "3 Apps Opened! - Halt"
        gblnSPAWN_DEBUG = True
        Stop
    End If

PROC_EXIT:
    Exit Function

PROC_ERR:
    If Err = 9 Then     ' Subscript out of range
        ReDim Preserve malngAccessHandles(1)
        malngAccessHandles(1) = mlngH
        Debug.Print , "malngAccessHandles(1) = " & malngAccessHandles(1)
        GetBounds = 1
    Else
        MsgBox "Erl=" & Erl & " Err=" & Err & " " & Err.Description, vbCritical, "GetBounds: Error"
        Resume PROC_EXIT
    End If

End Function