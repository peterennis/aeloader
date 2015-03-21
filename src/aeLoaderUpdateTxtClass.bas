Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

' Ref: Modules: Creating a reference to a Class in a Library database
' http://www.mvps.org/access/modules/mdl0034.htm

'********************************************************************************
' What:         A class to determine the update status of an Access
'               application database. Provides capabilities to work with
'               Access login, Windows network login or SQL Server login.
' Author:       (c) 1999 - 2015 Peter F. Ennis
'********************************************************************************

Private Declare PtrSafe Function GetComputerName Lib "kernel32.dll" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

' APPLICATION UPDATE GLOBAL VARIABLES
Private mstrAppName As String
Private mstrAppNewVersion As String
Private mastrUpdateUsers() As String
Private mstrUpdateText As String
Private mstrTheCurrentUser As String

' For SQL
Private mstrVersionDate As String
Private mblnUpdateUserSQL As String
Private mblnUpdateAllSQL As String
'

Public Property Let aeUpdateDebug(bln As Boolean)
' Allow Debug to be turned on outside of the class
    gfUpdateDebug = bln
End Property

Public Property Get aeUpdateDebug() As Boolean
    aeUpdateDebug = gfUpdateDebug
End Property

Private Sub aeDebugIt(strAbsLogFile As String, strData As String)
    Open strAbsLogFile For Append As #101
        Print #101, Date, "Now="; Format(Now(), "Short Time"), strData
    Close #101
End Sub

Private Sub aeOpenFrmLoaderUpdateNotes()
    If aeUpdateDebug Then
        DoCmd.OpenForm "frmLoaderUpdateNotes", , , , , acDialog, True
    Else
        DoCmd.OpenForm "frmLoaderUpdateNotes", , , , , acDialog
    End If
End Sub

Private Function GetUserList(strToParse As String) As String()
' What:     Function to get a list of users
' In:       A comma separated list of user names
' Out:      An array of user names

On Error GoTo Err_GetUserList

    Dim strToParseOriginal
    Dim astrUsers() As String
    ReDim astrUsers(10)
    Dim i As Integer
    Dim j As Integer

    j = 0
    strToParseOriginal = strToParse
    For i = 1 To Len(strToParseOriginal)
        'Debug.Print "j = " & j
        If InStr(1, strToParse, ",") = 0 Then
            astrUsers(j) = Trim(strToParse)
        Else
            astrUsers(j) = Trim(Mid(strToParse, 1, InStr(1, strToParse, ",") - 1))
        End If
        'Debug.Print astrUsers(j)
        If UCase(astrUsers(j)) = "ALL" Then
            ReDim astrUsers(1)
            astrUsers(0) = "ALL"
            GetUserList = astrUsers
            Debug.Print "astrUsers(0) = " & astrUsers(0)
            Exit Function
        End If
        '
        If InStr(1, strToParse, ",") <> 0 Then
            i = InStr(1, strToParse, ",") + 2
        End If
        'Debug.Print i
        strToParse = Trim(Mid(strToParse, i, Len(strToParse)))
        'Debug.Print strToParse
        'Debug.Print "Original Length = " & Len(strToParseOriginal) & vbCrLf & _
        '            "  Remain Length = " & Len(strToParse)
        If Len(strToParse) = Len(strToParseOriginal) Then
            ReDim Preserve astrUsers(j)
            GetUserList = astrUsers
            Exit For
        End If
        If Len(strToParse) = 0 Then
            ReDim Preserve astrUsers(j)
            GetUserList = astrUsers
            Exit For
        End If
        j = j + 1
    Next i
    
Exit_GetUserList:
    Exit Function

Err_GetUserList:
    MsgBox Err.Description & vbCrLf & vbCrLf & _
        "Probable format error in user update list.", _
        vbCritical, "DbLib: GetUserList " & Err
    Resume Exit_GetUserList

End Function

Public Function blnTheAppLoaderUpdateStatus() As Boolean

' What:         Function to test all requirenents for program update operation
' Author:       (c) Peter F. Ennis
' Created:      2/10/2000
' Passed in:    Application name to be updated, the server path, local path, update info file,
'               application version and application update filename as
'               strings by value. A debug flag by value.
' Returns:      True if system is to be updated.
' Last Mod:     02/22/2000
'               08/24/2004 adapt to use test mode from library
'                           called from frmUpdateNotes
'               09/12/2004 created as part of class
'               05/06/2005 aeDebugIt caused Invalid File Access Error #75.
'                           The log file had admin only permissions set from testing
'                           as admin. Quick fix - set permissions for everyone.
'                           Correct fix - app should write to log files with admin permissions
'                           regardless of who the user is.

On Error GoTo Err_blnTheAppLoaderUpdateStatus

    Dim strFileData As String
    Dim strAbsoluteFileName As String
    Dim strUpdateID As String
    Dim strUpdateID_Param As String
    Dim fAppNameCaptured As Boolean
    Dim intFileNum As Integer
    Dim intTemp As Integer
    Dim intTemp2 As Integer

    ' The Update Info File will be opened to determine how to do the update
    '
    ' Sample Contents:
    '
    'APP_NAME: DSFRC Intake
    'APP_NEW_VERSION: 4.0.1
    'APP_UPDATE_USER: pfe, dbuser:Station-131
    'APP_UPDATE_END:
    '

'?    If gstrAppName = "" Then
'?        MsgBox "Incorrect Application Name Setup" & vbCrLf & _
'?                "Please configure and use aeUpdateSetupClass", vbCritical, "aedb update library"
'?        Exit Function
'?    End If
    
1:    mstrUpdateText = ""                         ' Initialize update text string variable
2:    strAbsoluteFileName = gstrServerPath & gstrUpdateInfoFile
3:    'MsgBox "strAbsoluteFileName = " & strAbsoluteFileName
4:    If Not FileExists(strAbsoluteFileName) Then
5:        'MsgBox strAbsoluteFileName & " NOT FOUND!", vbCritical, "Application Update Function"
6:        blnTheAppLoaderUpdateStatus = False
7:        Exit Function
8:    Else
        'READ AND ASSIGN THE VALUES AND INFORMATION IN UPDATE INFO FILE
9:        blnTheAppLoaderUpdateStatus = True                        ' Initialize global update flag
10:        intFileNum = FreeFile                               ' Get available file number.

11:        Open strAbsoluteFileName For Input As intFileNum    ' Open to read file.
12:        Do While Not EOF(intFileNum)                        ' Check for end of file.
13:            Line Input #intFileNum, strFileData             ' Read line of data.
            'Debug.Print strFileData
14:            intTemp = InStr(1, strFileData, ":") - 1
            'Debug.Print Comment(strFileData)
15:            If (intTemp > 0) And Not Comment(strFileData) Then
16:                strUpdateID = Mid(strFileData, 1, intTemp)
17:                intTemp2 = Len(strFileData) - intTemp - 1
18:                strUpdateID_Param = Trim(Right(strFileData, intTemp2))
                'Debug.Print strUpdateID_Param
19:                If strUpdateID_Param = gstrAppName Then
20:                    fAppNameCaptured = True                 ' Application name found in update file
21:                End If
                'Debug.Print fAppNameCaptured
22:                If fAppNameCaptured Then
23:                    Select Case strUpdateID                 ' Data definition in Update.txt
                           Case "APP_NAME"
24:                            mstrAppName = strUpdateID_Param
25:                            Debug.Print "1> mstrAppName = " & strUpdateID_Param
                           Case "APP_NEW_VERSION"
26:                            mstrAppNewVersion = strUpdateID_Param
27:                            Debug.Print "2> mstrAppNewVersion = " & strUpdateID_Param & vbCrLf & _
                                        "   gstrAppCurrentVer = " & gstrAppCurrentVer
                            'IF NO CHANGE TO VERSION No. THEN DO NOT UPDATE
28:                            If (mstrAppNewVersion <= gstrAppCurrentVer) Then
29:                                If aeUpdateDebug Then
'MsgBox "gstrLocalPath = " & gstrLocalPath
'MsgBox "gstrDebugFile = " & gstrDebugFile
'MsgBox "gstrLocalPath & gstrDebugFile = " & gstrLocalPath & gstrDebugFile
'MsgBox "gstrDbLibVersion = " & gstrDbLibVersion
30:                                   aeDebugIt gstrLocalPath & gstrDebugFile, "gstrDbLibVersion = " & gstrDbLibVersion
'MsgBox "gstrAppCurrentVer = " & gstrAppCurrentVer
31:                                    aeDebugIt gstrLocalPath & gstrDebugFile, "gstrAppCurrentVer = " & gstrAppCurrentVer
32:                                    aeDebugIt gstrLocalPath & gstrDebugFile, "NOT UPDATING: No change to version number"
33:                                    blnTheAppLoaderUpdateStatus = False
34:                                    GoTo Exit_blnTheAppLoaderUpdateStatus
341:                                Else
342:                                    blnTheAppLoaderUpdateStatus = False
343:                                    GoTo Exit_blnTheAppLoaderUpdateStatus
35:                                End If
36:                            End If
                           Case "APP_NEW_FILE_VERSION"
361:                            mstrAppNewVersion = strUpdateID_Param
362:                            Debug.Print "2> mstrAppNewVersion = " & strUpdateID_Param & vbCrLf & _
                                        "   gstrAppCurrentVer = " & gstrAppCurrentVer
                            'IF NO CHANGE TO VERSION No. THEN DO NOT UPDATE
363:                            If (mstrAppNewVersion <= gstrAppCurrentVer) Then
364:                                If aeUpdateDebug Then
365:                                   aeDebugIt gstrLocalPath & gstrDebugFile, "gstrDbLibVersion = " & gstrDbLibVersion
366:                                    aeDebugIt gstrLocalPath & gstrDebugFile, "gstrAppCurrentVer = " & gstrAppCurrentVer
367:                                    aeDebugIt gstrLocalPath & gstrDebugFile, "NOT UPDATING: No change to version number"
368:                                    blnTheAppLoaderUpdateStatus = False
369:                                    GoTo Exit_blnTheAppLoaderUpdateStatus
370:                                Else
371:                                    blnTheAppLoaderUpdateStatus = False
372:                                    GoTo Exit_blnTheAppLoaderUpdateStatus
373:                                End If
374:                            End If
                           Case "APP_UPDATE_USER"
37:                            mastrUpdateUsers = GetUserList(UCase(strUpdateID_Param))
38:                            Debug.Print "3> mastrUpdateUsers = "
39:                            Dim i As Integer
40:                            For i = 0 To UBound(mastrUpdateUsers)
41:                                Debug.Print mastrUpdateUsers(i)
42:                            Next i
                            'IF NO DEFINED UPDATE USER THEN DO NOT UPDATE
43:                            If mastrUpdateUsers(0) = "NONE" Then
44:                                blnTheAppLoaderUpdateStatus = False
45:                                GoTo Exit_blnTheAppLoaderUpdateStatus
46:                            End If
                            'IF NOT ALL THEN TEST FOR LEGITIMATE GROUP NAME THEN TEST FOR LEGITIMATE USER NAME
47:                            If mastrUpdateUsers(0) <> "ALL" Then
48:                                mstrTheCurrentUser = UCase(gstrTheCurrentUser)
49:                                Debug.Print "4> mstrTheCurrentUser = " & mstrTheCurrentUser
                                ' CHECK ALL USERS
50:                                Dim j As Integer
51:                                For j = 0 To UBound(mastrUpdateUsers)
52:                                    If fTestUserName(mastrUpdateUsers(j)) <> mstrTheCurrentUser Then
53:                                        Debug.Print "fTestUserName(mastrUpdateUsers(j)) = " & fTestUserName(mastrUpdateUsers(j))
54:                                        blnTheAppLoaderUpdateStatus = False
55:                                    Else
56:                                        If InStr(1, mastrUpdateUsers(j), ":", 1) = 0 Then
57:                                            Debug.Print mastrUpdateUsers(j) & " WILL BE UPDATED."
58:                                            blnTheAppLoaderUpdateStatus = True
59:                                            Exit For
60:                                        Else
                                            ' Check for user at a particular machine
61:                                            If fTestComputerName(mastrUpdateUsers(j)) = aedblib_GetComputerName() Then
62:                                                Debug.Print mastrUpdateUsers(j) & " at computer " & _
                                                    aedblib_GetComputerName() & " WILL BE UPDATED."
63:                                                blnTheAppLoaderUpdateStatus = True
64:                                                Exit For
65:                                            Else
66:                                                blnTheAppLoaderUpdateStatus = False
67:                                            End If
68:                                        End If
69:                                    End If
70:                               Next j
71:                            End If
                           Case "APP_UPDATE_END"
72:                            Debug.Print "5> " & "APP_UPDATE_END"
73:                            fAppNameCaptured = False    ' Drop out of loop
74:                        Case Else
75:                    End Select
76:                End If
77:            End If
78:            If Not Comment(strFileData) And fAppNameCaptured Then        ' Create the update text string
79:                mstrUpdateText = mstrUpdateText & strFileData & vbCrLf
80:            End If
81:        Loop
82:    End If
    '
83:    gstrUpdateText = mstrUpdateText             ' Store value in frmUpdateNotes
84:    gstrAppNewVersion = mstrAppNewVersion
    '
85:    Close #intFileNum                   ' Close data file.

86:    If blnTheAppLoaderUpdateStatus Then
87:        If aeUpdateDebug Then
88:            aeDebugIt gstrLocalPath & gstrDebugFile, vbCrLf & _
            "aeloader Version = " & gconTHIS_APP_VERSION & vbCrLf & _
            "gstrAppName = " & gstrAppName & vbCrLf & _
            "gstrServerPath = " & gstrServerPath & vbCrLf & _
            "gstrLocalPath = " & gstrLocalPath & vbCrLf & _
            "gstrLocalLibPath = " & gstrLocalLibPath & vbCrLf & _
            "gstrUpdateInfoFile = " & gstrUpdateInfoFile & vbCrLf & _
            "gstrAppCurrentVer = " & gstrAppCurrentVer & vbCrLf & _
            "gstrUpdateAppFile = " & gstrUpdateAppFile & vbCrLf & _
            "gstrDebugFile = " & gstrDebugFile & vbCrLf & _
            "gfUpdateDebug = " & gfUpdateDebug & vbCrLf & _
            "_______________________________________________" & vbCrLf & _
            "mstrUpdateText:" & vbCrLf & _
            mstrUpdateText & vbCrLf & _
            "_______________________________________________"
89:            aeOpenFrmLoaderUpdateNotes
90:        Else
91:            aeOpenFrmLoaderUpdateNotes
92:        End If
93:        DoCmd.Quit
94:    End If

Exit_blnTheAppLoaderUpdateStatus:
    Exit Function

Err_blnTheAppLoaderUpdateStatus:
    MsgBox "Erl=" & Erl & " " & Err.Description, vbCritical, "Err_blnTheAppLoaderUpdateStatus Err=" & Err
    MsgBox "gstrAppName = " & gstrAppName & vbCrLf & _
            "gstrServerPath = " & gstrServerPath & vbCrLf & _
            "gstrLocalPath = " & gstrLocalPath & vbCrLf & _
            "gstrLocalLibPath = " & gstrLocalLibPath & vbCrLf & _
            "gstrUpdateInfoFile = " & gstrUpdateInfoFile & vbCrLf & _
            "gstrAppCurrentVer = " & gstrAppCurrentVer & vbCrLf & _
            "gstrUpdateAppFile = " & gstrUpdateAppFile & vbCrLf & _
            "gstrDebugFile = " & gstrDebugFile & vbCrLf & _
            "gfUpdateDebug = " & gfUpdateDebug & vbCrLf
    Resume Exit_blnTheAppLoaderUpdateStatus

End Function

Private Function fTestUserName(strUserName As String) As String

    Dim intSeparatorPos As Integer
    
'    strUser1 = "dbuser:Station-131"
'    intSeparatorPos1 = InStr(1, strUser1, ":", 1)
'    strUser2 = "dbuser"
'    intSeparatorPos2 = InStr(1, strUser2, ":", 1)
    
    intSeparatorPos = InStr(1, strUserName, ":", 1)
    If intSeparatorPos > 0 Then
        'Debug.Print "A: " & Mid(strUser1, 1, intSeparatorPos1 - 1)
        fTestUserName = Mid(strUserName, 1, intSeparatorPos - 1)
    Else
        'Debug.Print "B: " & Mid(strUserName, 1, intSeparatorPos2)
        fTestUserName = strUserName
    End If

End Function

Private Function fTestComputerName(strUserName As String) As String

    Dim intSeparatorPos As Integer
    
    intSeparatorPos = InStr(1, strUserName, ":", 1)
    Debug.Print intSeparatorPos
    If intSeparatorPos > 0 Then
        fTestComputerName = Mid(strUserName, intSeparatorPos + 1, Len(strUserName) - intSeparatorPos)
    Else
        fTestComputerName = "ThereIsNoComputerName"
    End If

End Function

Private Function aedblib_GetComputerName() As Variant
' Wrapper function for API GetComputerNameA routine

    Dim strComputerName As String
    Dim lngLength As Long
    Dim lngResult As Long
    
    ' Set up the buffer
    strComputerName = String$(255, 0)
    lngLength = 255
    ' Make the call
    lngResult = GetComputerName(strComputerName, lngLength)
    ' Clean up and assign the value
    aedblib_GetComputerName = Left(strComputerName, InStr(1, strComputerName, _
                                Chr(0)) - 1)
    
End Function