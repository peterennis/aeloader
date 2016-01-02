Option Compare Database
Option Explicit

'------------------- Declarations for getting the list of open windows
' Ref: http://support.microsoft.com/default.aspx?scid=kb;EN-US;168829
Public Const GW_CHILD = 5
Public Const GW_HWNDNEXT = 2
'
Public Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare PtrSafe Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare PtrSafe Function GetTopWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
'-------------------

' GLOBAL CONSTANTS
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
Public gstrTheServer As String
Public gstrTheWorkgroupFile As String
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
Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare PtrSafe Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const WM_CLOSE = &H10

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, _
        ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
'
Public Const SW_HIDE = 0
Public Const SW_SHOWNORMAL = 1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMAXIMIZED = 3

Private Declare PtrSafe Function apiShowWindow Lib "user32" Alias "ShowWindow" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Public Function GetTheUpdateFile() As Boolean

    Debug.Print "GetTheUpdateFile"

    On Error GoTo PROC_ERR

    Dim blnUpdate As Boolean
    Dim Setup As aeLoaderUpdateSetupClass
    Set Setup = New aeLoaderUpdateSetupClass

    ' Setup parameters
    Setup.aeUpdateDebug = True

    Dim strThePassThroughAppName As String
    Dim strThePassThroughAppVersion As String
    Debug.Print , "gintApp = " & gintApp
    gintApp = 6
    strThePassThroughAppName = gstrLocalPath & DLookup("gstrAppName", _
                            "aeLoaderParameters_Table", "ParameterID=" & gintApp)
    Debug.Print , "strThePassThroughAppName = " & strThePassThroughAppName
    strThePassThroughAppVersion = gstrLocalPath & DLookup("gstrAppFileName", _
                            "aeLoaderParameters_Table", "ParameterID=" & gintApp)
    Debug.Print , "strThePassThroughAppVersion = " & strThePassThroughAppVersion
    blnUpdate = Setup.aeUpdateSetup(strThePassThroughAppName, _
                            strThePassThroughAppVersion, aeWindowsNetworkLogin)

    Dim cls2 As aeLoaderUpdateTxtClass
    Set cls2 = New aeLoaderUpdateTxtClass

    cls2.aeUpdateDebug = True
    blnUpdate = cls2.blnTheAppLoaderUpdateStatus()
    Debug.Print , "cls2.blnTheAppLoaderUpdateStatus = " & blnUpdate
    MsgBox "cls2.blnTheAppLoaderUpdateStatus = " & blnUpdate

PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Err=" & Err & " (" & Err.Description & ") in procedure GetTheUpdateFile"
    Resume PROC_EXIT

End Function

Private Sub InitializeLoaderVariables()

        ' Setup the loader variables
        gstrTheAppExtension = DLookup("gstrAppExt", "aeLoaderParameters_Table", _
                            "ParameterID=" & gintApp)
        gstrTheApp = gstrLocalPath & DLookup("gstrAppFileName", "aeLoaderParameters_Table", _
                            "ParameterID=" & gintApp) & _
                            "." & gstrTheAppExtension
        gstrTheAppWindowName = DLookup("gstrAppWindowName", "aeLoaderParameters_Table", _
                            "ParameterID=" & gintApp)
        gstrLocalPath = DLookup("gstrLocalPath", "aeLoaderParameters_Table", _
                            "ParameterID=" & gintApp)
        gstrUpdateAppFile = DLookup("gstrUpdateAppFile", "aeLoaderParameters_Table", _
                            "ParameterID=" & gintApp)
        gstrTheServer = DLookup("gstrServerPath", "aeLoaderParameters_Table", _
                            "ParameterID=" & gintApp)
        gstrTheWorkgroupFile = DLookup("gstrTheWorkgroupFile", "aeLoaderParameters_Table", _
                            "ParameterID=" & gintApp)
        gstrTheWorkgroup = gstrTheServer & gstrTheWorkgroupFile
        gstrLogonMdb = DLookup("gstrLogonMdb", "aeLoaderParameters_Table", _
                            "ParameterID=" & gintApp)
        gstrPasswordMdb = DLookup("gstrPasswordMdb", "aeLoaderParameters_Table", _
                            "ParameterID=" & gintApp)
        gstrDbLibName = DLookup("gstrDbLibName", "aeLoaderParameters_Table", _
                            "ParameterID=" & gintApp)
        ' Updates will occur in the application based on the version.
        ' A corresponding library can be called e.g. adaeptdblib.mda.v425
        ' copied across and renamed to adaeptdblib.mda.upd locally.
        gstrLocalLibPath = DLookup("gstrLocalLibPath", "aeLoaderParameters_Table", _
                            "ParameterID=" & gintApp)
        'MsgBox "gstrTheAppWindowName = " & gstrTheAppWindowName & vbCrLf & _
            "gstrLocalPath = " & gstrLocalPath & vbCrLf & _
            "gstrUpdateAppFile = " & gstrUpdateAppFile & vbCrLf & _
            "gstrTheAppExtension = " & gstrTheAppExtension & vbCrLf & _
            "gstrTheApp = " & gstrTheApp & vbCrLf & _
            "gstrLogonMdb = " & gstrLogonMdb & vbCrLf & _
            "gstrPasswordMdb = " & gstrPasswordMdb & vbCrLf & _
            "gstrLocalLibPath = " & gstrLocalLibPath & vbCrLf & _
            "gstrDbLibName = " & gstrDbLibName & vbCrLf & _
            "gstrPassThrough = " & gstrPassThrough & vbCrLf & _
            "gstrTheWorkgroupFile = " & gstrTheWorkgroupFile & vbCrLf & _
            "gstrTheServer = " & gstrTheServer & vbCrLf & _
            "gstrTheWorkgroup = " & gstrTheWorkgroup

End Sub

Public Function StartApp() As Boolean

    Debug.Print "StartApp"

    On Error GoTo PROC_ERR

    gstrPassThrough = Nz(DLookup("gstrPassThrough", "aeLoaderParameters_Table", "ParameterID=" & gintApp))
    Debug.Print , "gintApp = " & gintApp
    Debug.Print , "gstrPassThrough = " & gstrPassThrough

    If gstrPassThrough = "PassThrough" Then

        Dim blnUpdate As Boolean
        Dim cls1 As aeLoaderUpdateSetupClass
        Set cls1 = New aeLoaderUpdateSetupClass

        ' Setup parameters
        cls1.aeUpdateDebug = True

        Dim strThePassThroughAppName As String
        Dim strThePassThroughAppVersion As String
        strThePassThroughAppName = gstrLocalPath & DLookup("gstrAppName", _
                            "aeLoaderParameters_Table", "ParameterID=" & gintApp)
        strThePassThroughAppVersion = gstrLocalPath & DLookup("gstrAppFileName", _
                            "aeLoaderParameters_Table", "ParameterID=" & gintApp)
        blnUpdate = cls1.aeUpdateSetup(strThePassThroughAppName, _
                            strThePassThroughAppVersion, aeWindowsNetworkLogin)

        Dim cls2 As aeLoaderUpdateTxtClass
        Set cls2 = New aeLoaderUpdateTxtClass

        cls2.aeUpdateDebug = True
        blnUpdate = cls2.blnTheAppLoaderUpdateStatus()
        Debug.Print , "cls2.blnTheAppLoaderUpdateStatus = " & blnUpdate
        MsgBox "cls2.blnTheAppLoaderUpdateStatus = " & blnUpdate
        '
        ' Shutdown the app if it is already open
        'MsgBox "gstrTheAppWindowName = " & gstrTheAppWindowName, vbInformation, gconTHIS_APP_NAME & ": StartApp"
        Debug.Print "gstrTheAppWindowName = " & gstrTheAppWindowName, vbInformation, gconTHIS_APP_NAME & ": StartApp"
        ShutDownApplication (gstrTheAppWindowName)
        '
        StartApp = aeLoaderPassThroughApp(gstrLocalPath, gstrLoaderUpdateAppFile)
        DoCmd.Quit
        Exit Function

    Else

        ' Minimize the Access window
        Debug.Print , "Minimizing the Access window"
        ShowWindow Application.hWndAccessApp, SW_SHOWMINIMIZED

        InitializeLoaderVariables

        ' Shutdown the app if it is already open
        ShutDownApplication (gstrTheAppWindowName)

        ' Update to new library
        '''If gstrDbLibName <> "NONE" Then InstallNewLibrary

        Debug.Print , "StartApp: gstrLocalPath & gstrUpdateAppFile = " & gstrLocalPath & gstrUpdateAppFile
        StartApp = aeLoaderApp(gstrLocalPath & gstrUpdateAppFile)

        DoCmd.Restore
        DoCmd.Quit
    End If

PROC_EXIT:
    Exit Function

PROC_ERR:
    Select Case Err
'        Case 58
'            ' OLD app file exists
'            Kill Mid(strAbsAppName, 1, Len(strAbsAppName) - 3) & "OLD"
'            Resume
        Case 75
            ' Path/File access error: If app is open it takes time to be
            ' shut down so try again
            Delay 1
            Resume
        Case Else
            MsgBox "Erl=" & Erl & " Err=" & Err & " " & Err.Description, vbCritical, "StartApp: " & gconTHIS_APP_NAME
    End Select
    Resume PROC_EXIT

End Function

Public Function aeLoaderPassThroughApp(strPath As String, strFileName As String) As Boolean
' What:         Load the selected pass through application
' Author:       Peter F. Ennis
' Created:      9/13/2003
' Passed in:    Absolute application file name as a string
' Returns:      True if successful
' Last Mod:

    Debug.Print "aeLoaderPassThroughApp"

    On Error GoTo PROC_ERR

    If FileExists(strPath & strFileName) Then
        'MsgBox strPath & strFileName & " FOUND." & vbCrLf & _
            "WRITE CODE TO KILL OLD APPS", vbInformation, "aeLoaderPassThroughApp: " & gconTHIS_APP_NAME
         Debug.Print , "aeLoaderPassThroughApp: strPath = " & strPath
         Debug.Print , "aeLoaderPassThroughApp: strFileName = " & strFileName
        KillOldApps strPath, strFileName
    End If

    Do
        OpenNotSecured strPath & strFileName

        If gblnSPAWN_DEBUG Then
            Dim i As Integer
            i = MsgBox("L6 gblnSPAWN_DEBUG", vbYesNo, "Test Break")
            If i = vbYes Then
                Exit Function
            Else
            End If
        End If

        DoEvents
    Loop Until WindowIsOpen(gstrTheAppWindowName)

    aeLoaderPassThroughApp = True

PROC_EXIT:
    Exit Function

PROC_ERR:
    'MsgBox Err & " " & Err.Description, vbCritical, "aeLoaderPassThroughApp: " & gconTHIS_APP_NAME
    Select Case Err
        Case 75
            ' Path/File access error: If app is open it takes time to be
            ' shut down so try again
            Delay 1
            Resume
        Case Else
            MsgBox "Erl:" & Erl & " Error# " & Err & " " & Err.Description & vbCrLf & _
                    "strPath = " & strPath & _
                    "strFileName = " & strFileName, vbCritical, "aeLoaderPassThroughApp: " & gconTHIS_APP_NAME
    End Select
    aeLoaderPassThroughApp = False
    Resume PROC_EXIT

End Function

Public Sub KillOldApps(strPath As String, strFileName As String)

    Debug.Print "KillOldApps"
    Debug.Print , "strPath = " & strPath
    Debug.Print , "strFileName = " & strFileName

    On Error GoTo PROC_ERR

    Dim strFName As String
    Dim strFilePattern As String
    
    strFilePattern = Left(strFileName, InStr(strFileName, gstrTheAppSeparatorChar))
    Debug.Print , "strFilePattern = " & strFilePattern
    
    ' Display the names in strPath that represent the application to be started
    strFName = Dir(strPath & strFilePattern & "*")    ' Retrieve the first entry.
    Do While strFName <> ""    ' Start the loop.
         If strFName <> strFileName Then
             Debug.Print , "Found: " & strFName
             Kill strPath & strFName
        Else
            Debug.Print , "APP TO LOAD: " & strFName
        End If
        strFName = Dir    ' Get next entry.
    Loop
    'Stop
      ' Make copy of app bmp startup file
      If FileExists(strPath & gstrAppCmdName & ".bmp") Then
            'MsgBox "Creating App bmp File"
            FileCopy strPath & gstrAppCmdName & ".bmp", strPath & Mid(strFileName, 1, Len(strFileName) - 4) & ".bmp"
      End If

PROC_EXIT:
Exit Sub

PROC_ERR:
    MsgBox "Erl=" & Erl & " Err=" & Err & " " & Err.Description, vbCritical, "KillOldApps: " & gconTHIS_APP_NAME
    Resume Next

End Sub

Private Function aeLoaderApp(strAbsAppName As String) As Boolean
' What:         Load the selected application
' Author:       Peter F. Ennis
' Created:      8/2004
' Passed in:    Absolute application file name as a string
' Returns:      True if successful
' Last Mod:     See changes on GitHub - https://github.com/peterennis/aeloader

    Debug.Print "aeLoaderApp"
    Debug.Print , "strAbsAppName = " & strAbsAppName

    On Error GoTo PROC_ERR

    If FileExists(strAbsAppName) Then
        ' Rename the old app file
        Name Mid(strAbsAppName, 1, Len(strAbsAppName) - 3) & gstrTheAppExtension _
                As Mid(strAbsAppName, 1, Len(strAbsAppName) - 3) & "OLD"
        ' Rename the update app file
        Name strAbsAppName As Mid(strAbsAppName, 1, _
                Len(strAbsAppName) - 3) & gstrTheAppExtension
    Else
        MsgBox "Update file'" & strAbsAppName & "' not found!", vbCritical, "aeLoaderApp: " & gconTHIS_APP_NAME
        'Stop
        DoCmd.Quit
    End If

    Debug.Print , "aeLoaderApp: gstrTheWorkgroupFile = " & gstrTheWorkgroupFile
    If gstrTheWorkgroupFile <> "OBSOLETE" Then
        Debug.Print , "aeLoaderApp: Opening database with a secured workgroup"
        Do
            OpenSecured gstrTheApp, gstrTheWorkgroup, gstrLogonMdb, gstrPasswordMdb

            If gblnSPAWN_DEBUG Then
                Dim i As Integer
                i = MsgBox("L6", vbYesNo, "Test Break")
                If i = vbYes Then
                    Exit Function
                Else
                End If
            End If

            DoEvents
        Loop Until WindowIsOpen(gstrTheAppWindowName)
    Else
        Debug.Print , "aeLoaderApp: Opening normal database"
        Debug.Print , "aeLoaderApp: gstrLocalPath = " & gstrLocalPath
        Debug.Print , "aeLoaderApp: gstrTheApp = " & gstrTheApp
        'Stop
        Do
            'MsgBox "gstrTheAppWindowName = " & gstrTheAppWindowName, vbInformation, gconTHIS_APP_NAME
            'MsgBox "gstrLocalPath & gstrTheApp = " & gstrLocalPath & gstrTheApp, vbInformation, gconTHIS_APP_NAME
            OpenNotSecured gstrLocalPath & gstrTheApp

            If gblnSPAWN_DEBUG Then
                Dim j As Integer
                j = MsgBox("L62", vbYesNo, "Test Break")
                If j = vbYes Then
                    Exit Function
                Else
                End If
            End If

            DoEvents
        Loop Until WindowIsOpen(gstrTheAppWindowName)
    End If

    'MsgBox WindowIsOpen("The Window Title")
    'MaximizeTheWindow WindowIsOpen("The Window Title"), "The Window Title"

    aeLoaderApp = True

PROC_EXIT:
    Exit Function

PROC_ERR:
    'MsgBox Err & " " & Err.Description, vbCritical, "aeLoaderApp: " & gconTHIS_APP_NAME
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
            MsgBox "Erl:" & Erl & " Error# " & Err & " " & Err.Description, vbCritical, gconTHIS_APP_NAME & ": aeLoaderApp"
    End Select
    aeLoaderApp = False
    Resume PROC_EXIT

End Function

Public Function FileExists(strAbsFileName As String) As Boolean
' What:         Test for existence of a file.
' Author:       Peter F. Ennis
' Created:      11/1998
' Passed in:    Absolute file name as a string
' Returns:      True
' Last Mod:     07/30/99
'               08/27/2004 use gconTHIS_APP_NAME in messages

    Debug.Print "FileExists"
    Debug.Print , "strAbsFileName = " & strAbsFileName

    On Error GoTo PROC_ERR

    Dim strSubName As String
    strSubName = "FileExists"
    Dim strMessage As String

    FileExists = (Dir(strAbsFileName) <> "")
    Debug.Print , FileExists

PROC_EXIT:
    Exit Function

PROC_ERR:
    Select Case Err
        Case 53, 62
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
    Resume PROC_EXIT

End Function

Public Function ShutDownApplication(ByVal strApplicationName As String) As Boolean
' Ref: http://www.a1vbcode.com/app.asp?ID=479

    Debug.Print "ShutDownApplication"

    Dim hWnd As Long
    Dim Result As Long
    hWnd = FindWindow(vbNullString, strApplicationName)
    Debug.Print , "hWnd = " & hWnd
    Debug.Print , "strApplicationName = " & strApplicationName
    If hWnd <> 0 Then
        Result = PostMessage(hWnd, WM_CLOSE, 0&, 0&)
        'MsgBox "The application window was found for shutdown."
        ShutDownApplication = True
        ' If the app will not shutdown but give a Quit message then close the loader
        If gstrTheAppWindowName = "THIS APP WILL NOT SHUTDOWN" Then
            DoCmd.Quit
        End If
    Else
        'MsgBox "The application window " & _
            strApplicationName & " was not found.", vbInformation, _
            gconTHIS_APP_NAME & ": ShutDownApplication"
        Debug.Print , "The application window '" & strApplicationName & "' was not found"
        'Stop
        'DoCmd.Quit
        ShutDownApplication = False
    End If

End Function

Private Function WindowIsOpen(ByVal strWindowTitle As String) As Long

    Debug.Print "WindowIsOpen"

    Dim hWnd As Long
    Dim Result As Long
    hWnd = FindWindow(vbNullString, strWindowTitle)
    Debug.Print , "hWnd = " & hWnd
    If hWnd <> 0 Then
        WindowIsOpen = hWnd
    Else
        WindowIsOpen = 0
    End If

End Function

Private Sub MaximizeTheWindow(hWnd As Long, ByVal strWindowTitle As String)
'Ref: http://www.digital-inn.de/archive/index.php/t-15364.html

    Debug.Print "MaximizeTheWindow"

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
    
    Debug.Print "OpenSecured"
    
    Dim objAccess As Object
    Dim cmd As String
    
    On Error Resume Next
    Set objAccess = GetObject(, "Access.Application")
    If Err = 0 Then 'an instance of Access is open
        If IsMissing(varUser) Then varUser = "Admin"
        
' ******** EXAMPLE ********
'        cmd = """C:\Program Files\Microsoft Office\Office\MSAccess.exe""" & " " & _
'                 """C:\The\Database\SQL 2000 Front End A2K.mdb""" & " " & _
'                 "/wrkgrp" & " " & _
'                 "\\The\Server\TheWorkgroup.MDW" & " " & _
'                 "/cmd " & _
'                 """NOSPLASHFORM"""
'        MsgBox cmd
' **************************

' 12.0 = Access 2007
' 14.0 = Access 2010
' 15.0 = Access 2013
' 16.0 = Access 2016

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
        bln = IsRunning("access")
        Do 'Wait for shelled process to finish.
            Err = 0
            Set objAccess = GetObject(, "Access.Application")
        Loop While Err <> 0
    End If

End Sub

Public Function Is64bit() As Boolean
    Is64bit = Len(Environ("ProgramW6432")) > 0
End Function

Private Sub OpenNotSecured(strTheApp As String)

    Debug.Print "OpenNotSecured"

    Dim objAccess As Object
    Dim cmd As String
    Dim str64Bit As String
    
    On Error Resume Next
    Set objAccess = GetObject(, "Access.Application")

    If Is64bit Then
        str64Bit = " (x86)"
    Else
        str64Bit = ""
    End If
            
    If GetAccessVersion = "9.0" Then        ' Access 2000
        cmd = """" & "C:\Program Files" & str64Bit & "\Microsoft Office\Office\MSAccess.exe" & """" & " "
    ElseIf GetAccessVersion = "11.0" Then   ' Access 2003
        cmd = """" & "C:\Program Files" & str64Bit & "\Microsoft Office\Office11\MSAccess.exe" & """" & " "
    ElseIf GetAccessVersion = "12.0" Then   ' Access 2007
        cmd = """" & "C:\Program Files" & str64Bit & "\Microsoft Office\Office12\MSAccess.exe" & """" & " "
    ElseIf GetAccessVersion = "14.0" Then   ' Access 2010
        cmd = """" & "C:\Program Files" & str64Bit & "\Microsoft Office\Office14\MSAccess.exe" & """" & " "
    ElseIf GetAccessVersion = "15.0" Then   ' Access 2013
        cmd = """" & "C:\Program Files" & str64Bit & "\Microsoft Office\Office15\MSAccess.exe" & """" & " "
    ElseIf GetAccessVersion = "16.0" Then   ' Access 2016
        cmd = """" & "C:\Program Files" & str64Bit & "\Microsoft Office\Office16\MSAccess.exe" & """" & " "
    End If

    cmd = cmd & """" & strTheApp & """"
    Debug.Print , "cmd = " & cmd
    'MsgBox cmd, vbInformation, gconTHIS_APP_NAME & ": OpenNotSecured"
    '
    Shell pathname:=cmd, windowstyle:=vbMaximizedFocus
    Dim bln As Boolean
    bln = IsRunning("access")
    Do 'Wait for shelled process to finish.
        Err = 0
        Set objAccess = GetObject(, "Access.Application")
    Loop While Err <> 0

End Sub

Public Sub Delay(pdblSeconds As Double)
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

Private Function DoAccessWindow(nCmdShow As Long)
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

Private Function GetAccessVersion() As String
' Ref: http://www.blueclaw-db.com/get_access_version_number.htm
' To determine the version of Microsoft Access used to open this application.
' 8.0 = Access 97
' 9.0 = Access 2000
' 10.0 = Access 2002(XP)
' 11.0 = Access 2003
' Ref: http://en.wikipedia.org/wiki/Microsoft_Access
' 12.0 = Access 2007
' 14.0 = Access 2010
' 15.0 = Access 2013
' 16.0 = Access 2016

    GetAccessVersion = SysCmd(acSysCmdAccessVer)

End Function

Public Function aeGetTheAppID() As Integer
' Ref: http://www.microsoft.com/en-us/download/details.aspx?id=19494

    Dim intAppID As Integer

    'MsgBox "aeGetTheAppID: Command = " & Command, vbInformation, gconTHIS_APP_NAME
    gstrAppCmdName = Command
    If IsNull(gstrAppCmdName) Or gstrAppCmdName = vbNullString Then
        MsgBox "No Command parameter found." & vbCrLf & _
                "Did you start the loader from a shortcut?", vbCritical, gconTHIS_APP_NAME & ": aeGetTheAppID"
        'Stop
        DoCmd.Quit
        Exit Function
    End If

    intAppID = Nz(DLookup("[ParameterID]", "[aeLoaderParameters_Table]", _
                        "[gstrAppName] = '" & gstrAppCmdName & "'"))
    'MsgBox "aeGetTheAppID: intAppID = " & intAppID & vbCrLf & _
                "gstrAppCmdName = " & gstrAppCmdName, vbInformation, gconTHIS_APP_NAME

    If intAppID = 0 Then
        MsgBox "aeGetTheAppID: Invalid Access Command Line Parameter!" & vbCrLf & vbCrLf & _
                "Command = " & "'" & Command & "'", vbCritical, gconTHIS_APP_NAME
        DoCmd.Restore
        DoCmd.Quit acQuitSaveNone
        Exit Function
    End If

    aeGetTheAppID = intAppID
    
End Function

Public Function Comment(strComment As String) As Boolean
' What:         This function returns true if a string is a comment i.e. if
'               the first character in the first line is ' OR ;
' Author:       Peter F. Ennis
' Date          11/98
' Parameter:    String to be verified as a comment
' Returns:      True if the string is a comment and false otherwise
' Last Mod:     7/30/99

    On Error GoTo PROC_ERR

    Comment = False
    If ((Mid$(strComment, 1, 1) = "'") Or (Mid$(strComment, 1, 1) = ";")) Then
        Comment = True
    End If

PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Comment Error " & Err & ": " & Error$, vbCritical, "aedb"
    Resume PROC_EXIT

End Function

Public Function MoveToCenter()

    Dim bln As Boolean
    Dim cls As aeLoaderMoveSizeClass
    Set cls = New aeLoaderMoveSizeClass
    ' Setup parameters
    cls.aeMoveSizeCenter = True

End Function