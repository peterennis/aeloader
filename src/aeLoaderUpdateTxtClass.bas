Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

' Ref: Modules: Creating a reference to a Class in a Library database
' http://www.mvps.org/access/modules/mdl0034.htm

' 09/12/2004 Reference to class created in library
' 09/01/2005 Import from adaeptdblib.mda and modify with SetupClass to use
'               PASS_THROUGH for updating based on APP_NEW_FILE_VERSION change
' 05/09/2007 Allow check for user at a machine = APP_UPDATE_USER: pfe, dbuser:Station-131
'

'********************************************************************************
' What:         A class to determine the update status of an Access
'               application database. Provides capabilities to work with
'               Access login, Windows network login or SQL Server login.
' Author:       (c) 1999 - 2007 Peter F. Ennis
'********************************************************************************

Private Declare Function GetComputerName Lib "kernel32.dll" Alias "GetComputerNameA" _
    (ByVal lpBuffer As String, nSize As Long) As Long

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

Private Sub aeOpenFrmUpdateNotes()
    If aeUpdateDebug Then
        DoCmd.OpenForm "frmUpdateNotes", , , , , acDialog, True
    Else
        DoCmd.OpenForm "frmUpdateNotes", , , , , acDialog
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
    mstrUpdateText = ""                         ' Initialize update text string variable
    strAbsoluteFileName = gstrServerPath & gstrUpdateInfoFile
    'MsgBox "strAbsoluteFileName = " & strAbsoluteFileName
    If Not FileExists(strAbsoluteFileName) Then
        'MsgBox strAbsoluteFileName & " NOT FOUND!", vbCritical, "Application Update Function"
        blnTheAppLoaderUpdateStatus = False
        Exit Function
    Else
        'READ AND ASSIGN THE VALUES AND INFORMATION IN UPDATE INFO FILE
        blnTheAppLoaderUpdateStatus = True                        ' Initialize global update flag
        intFileNum = FreeFile                               ' Get available file number.
'MsgBox "1"
        Open strAbsoluteFileName For Input As intFileNum    ' Open to read file.
'MsgBox "2"
        Do While Not EOF(intFileNum)                        ' Check for end of file.
'MsgBox "3"
            Line Input #intFileNum, strFileData             ' Read line of data.
'MsgBox "4"
            'Debug.Print strFileData
            intTemp = InStr(1, strFileData, ":") - 1
            'Debug.Print Comment(strFileData)
'MsgBox "5"
            If (intTemp > 0) And Not Comment(strFileData) Then
'MsgBox "5.1"
                strUpdateID = Mid(strFileData, 1, intTemp)
'MsgBox "5.2"
                intTemp2 = Len(strFileData) - intTemp - 1
'MsgBox "5.3"
                strUpdateID_Param = Trim(right(strFileData, intTemp2))
                'Debug.Print strUpdateID_Param
'MsgBox "5.4"
                If strUpdateID_Param = gstrAppName Then
                    fAppNameCaptured = True                 ' Application name found in update file
                End If
'MsgBox "5.5"
                'Debug.Print fAppNameCaptured
                If fAppNameCaptured Then
                    Select Case strUpdateID                 ' Data definition in Update.txt
                        Case "APP_NAME"
'MsgBox "5.6"
                            mstrAppName = strUpdateID_Param
                            Debug.Print "1> mstrAppName = " & strUpdateID_Param
                        Case "APP_NEW_VERSION"
'MsgBox "5.7.1"
                            mstrAppNewVersion = strUpdateID_Param
'MsgBox "5.7.2"
                            Debug.Print "2> mstrAppNewVersion = " & strUpdateID_Param & vbCrLf & _
                                        "   gstrAppCurrentVer = " & gstrAppCurrentVer
                            'IF NO CHANGE TO VERSION No. THEN DO NOT UPDATE
'MsgBox "5.7.3"
                            If (mstrAppNewVersion <= gstrAppCurrentVer) Then
'MsgBox "5.7.4"
                                If aeUpdateDebug Then
'MsgBox "5.7.5"
'MsgBox "gstrLocalPath = " & gstrLocalPath
'MsgBox "gstrDebugFile = " & gstrDebugFile
'MsgBox "gstrLocalPath & gstrDebugFile = " & gstrLocalPath & gstrDebugFile
'MsgBox "gstrDbLibVersion = " & gstrDbLibVersion
                                    aeDebugIt gstrLocalPath & gstrDebugFile, "gstrDbLibVersion = " & gstrDbLibVersion
'MsgBox "5.7.5.1"
'MsgBox "gstrAppCurrentVer = " & gstrAppCurrentVer
                                    aeDebugIt gstrLocalPath & gstrDebugFile, "gstrAppCurrentVer = " & gstrAppCurrentVer
'MsgBox "5.7.5.2"
                                    aeDebugIt gstrLocalPath & gstrDebugFile, "NOT UPDATING: No change to version number"
'MsgBox "5.7.5.3"
                                    blnTheAppLoaderUpdateStatus = False
'MsgBox "5.7.5.4"
                                    GoTo Exit_blnTheAppLoaderUpdateStatus
                                End If
'MsgBox "5.7.6"
                            End If
                        Case "APP_UPDATE_USER"
'MsgBox "5.8"
                            mastrUpdateUsers = GetUserList(UCase(strUpdateID_Param))
                            Debug.Print "3> mastrUpdateUsers = "
                            Dim i As Integer
                            For i = 0 To UBound(mastrUpdateUsers)
                                Debug.Print mastrUpdateUsers(i)
                            Next i
                            'IF NO DEFINED UPDATE USER THEN DO NOT UPDATE
                            If mastrUpdateUsers(0) = "NONE" Then
                                blnTheAppLoaderUpdateStatus = False
                                GoTo Exit_blnTheAppLoaderUpdateStatus
                            End If
                            'IF NOT ALL THEN TEST FOR LEGITIMATE GROUP NAME THEN TEST FOR LEGITIMATE USER NAME
                            If mastrUpdateUsers(0) <> "ALL" Then
                                mstrTheCurrentUser = UCase(gstrTheCurrentUser)
                                Debug.Print "4> mstrTheCurrentUser = " & mstrTheCurrentUser
                                ' CHECK ALL USERS
                                Dim j As Integer
                                For j = 0 To UBound(mastrUpdateUsers)
                                    If fTestUserName(mastrUpdateUsers(j)) <> mstrTheCurrentUser Then
                                        Debug.Print "fTestUserName(mastrUpdateUsers(j)) = " & fTestUserName(mastrUpdateUsers(j))
                                        blnTheAppLoaderUpdateStatus = False
                                    Else
                                        If InStr(1, mastrUpdateUsers(j), ":", 1) = 0 Then
                                            Debug.Print mastrUpdateUsers(j) & " WILL BE UPDATED."
                                            blnTheAppLoaderUpdateStatus = True
                                            Exit For
                                        Else
                                            ' Check for user at a particular machine
                                            If fTestComputerName(mastrUpdateUsers(j)) = aedblib_GetComputerName() Then
                                                Debug.Print mastrUpdateUsers(j) & " at computer " & _
                                                    aedblib_GetComputerName() & " WILL BE UPDATED."
                                                blnTheAppLoaderUpdateStatus = True
                                                Exit For
                                            Else
                                                blnTheAppLoaderUpdateStatus = False
                                            End If
                                        End If
                                    End If
                               Next j
                            End If
                        Case "APP_UPDATE_END"
'MsgBox "5.9"
                            Debug.Print "5> " & "APP_UPDATE_END"
                            fAppNameCaptured = False    ' Drop out of loop
                        Case Else
                    End Select
                End If
            End If
'MsgBox "6"
            If Not Comment(strFileData) And fAppNameCaptured Then        ' Create the update text string
                mstrUpdateText = mstrUpdateText & strFileData & vbCrLf
            End If
        Loop
'MsgBox "7"
    End If
    '
    gstrUpdateText = mstrUpdateText             ' Store value in frmUpdateNotes
    gstrAppNewVersion = mstrAppNewVersion
    '
    If blnTheAppLoaderUpdateStatus Then
        If aeUpdateDebug Then
            aeOpenFrmUpdateNotes
            aeDebugIt gstrLocalPath & gstrDebugFile, vbCrLf & _
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
            Close #intFileNum                   ' Close data file.
        Else
            aeOpenFrmUpdateNotes
        End If
        DoCmd.Quit
    End If
    
Exit_blnTheAppLoaderUpdateStatus:
    Exit Function

Err_blnTheAppLoaderUpdateStatus:
    MsgBox "DbLib blnTheAppLoaderUpdateStatus Error # " & Err.Number & ": " & Err.Description
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