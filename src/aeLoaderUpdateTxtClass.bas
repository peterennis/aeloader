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

'********************************************************************************
' What:         A class to determine the update status of any
'               application based on filename changes.
' Author:       (c) 1999 - 2005 Peter F. Ennis
'********************************************************************************
    
' APPLICATION UPDATE GLOBAL VARIABLES
Private mstrAppName As String
Private mstrAppNewFileVersion As String
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
        vbCritical, "aeLoaderUpdateTxtClass: GetUserList " & Err
    Resume Exit_GetUserList

End Function

Public Function blnTheAppLoaderUpdateStatus() As Boolean

' What:         Function to test all requirenents for program update operation
' Author:       (c) Peter F. Ennis
' Created:      2/10/2000
' Passed in:
' Returns:      True if system is to be updated.
' Last Mod:     02/22/2000
'               08/24/2004 modify to use test mode from library
'                           called from frmLoaderUpdateNotes
'               09/12/2004 created as part of class
'               05/06/2005 aeDebugIt caused Invalid File Access Error #75.
'                           The log file had admin only permissions set from testing
'                           as admin. Quick fix - set permissions for everyone.
'                           Correct fix - app should write to log files with admin permissions
'                           regardless of who the user is.
'               09/06/2005 Modify for aeLoader and Noho


On Error GoTo Err_blnTheAppLoaderUpdateStatus

    Dim strFileData As String
    Dim strAbsoluteFileName As String
    Dim strUpdateID As String
    Dim strUpdateID_Param As String
    Dim fAppNameCaptured As Boolean
    Dim intFileNum As Integer
    Dim intTemp As Integer
    Dim intTemp2 As Integer
    Dim intLoop As Integer

    ' The Update Info File will be opened to determine how to do the update
    '
    ' Sample Contents:
    '
    'APP_NAME: The Application
    '#APP_NEW_VERSION:
    'APP_NEW_FILE_VERSION: TheNewFile-v100.mde
    '# Use ALL for updating all users
    'APP_UPDATE_USER: pfe
    'APP_UPDATE_END:
    '

     'MsgBox "aeLoaderUpdateTxtClass: blnTheAppLoaderUpdateStatus: aeUpdateDebug = " & aeUpdateDebug
1    If gstrTheApp = "" Then
2        MsgBox "Incorrect Application Name Setup" & vbCrLf & _
                "Please configure and use aeLoaderUpdateSetupClass", vbCritical, _
                "aeLoaderUpdateTxtClass: blnTheAppLoaderUpdateStatus"
3        Exit Function
4    End If
5    mstrUpdateText = ""    ' Initialize update text string variable
6    strAbsoluteFileName = gstrServerPath & gstrUpdateInfoFile
     'MsgBox "aeLoaderUpdateTxtClass: blnTheAppLoaderUpdateStatus: strAbsoluteFileName = " & strAbsoluteFileName
601  intLoop = 1
7    If Not FileExists(strAbsoluteFileName) Then
        'MsgBox "strAbsoluteFileName & " NOT FOUND!", vbCritical, _
            "aeLoaderUpdateTxtClass: blnTheAppLoaderUpdateStatus"
8        blnTheAppLoaderUpdateStatus = False
9        Exit Function
10    Else
          'READ AND ASSIGN THE VALUES AND INFORMATION IN UPDATE INFO FILE
11        blnTheAppLoaderUpdateStatus = True                  ' Initialize global update flag
12        intFileNum = FreeFile                               ' Get available file number.
13        Open strAbsoluteFileName For Input As intFileNum    ' Open to read file.
14        Do While Not EOF(intFileNum)                        ' Check for end of file.
15            Line Input #intFileNum, strFileData             ' Read line of data.
              'Debug.Print strFileData
16            intTemp = InStr(1, strFileData, ":") - 1
              'Debug.Print Comment(strFileData)
17            If (intTemp > 0) And Not Comment(strFileData) Then
18                strUpdateID = Mid(strFileData, 1, intTemp)
19                intTemp2 = Len(strFileData) - intTemp - 1
20                strUpdateID_Param = Trim(right(strFileData, intTemp2))
                  Debug.Print fAppNameCaptured, strUpdateID, strUpdateID_Param, gstrTheApp
21                If strUpdateID = "APP_NEW_FILE_VERSION" And Not FileExists(gstrLocalPath & strUpdateID_Param) Then
22                    fAppNameCaptured = True                 ' Application name found in update file
23                End If
231               If strUpdateID = "APP_NEW_FILE_VERSION" And FileExists(gstrLocalPath & strUpdateID_Param) Then
232                      gstrLoaderUpdateAppFile = strUpdateID_Param
233                   'MsgBox "233 gstrLoaderUpdateAppFile = " & gstrLoaderUpdateAppFile, vbInformation, _
                          "aeLoaderUpdateTxtClass: blnTheAppLoaderUpdateStatus"
234                   blnTheAppLoaderUpdateStatus = False
235                   GoTo Exit_blnTheAppLoaderUpdateStatus
236               End If
                  'Debug.Print fAppNameCaptured
24                If fAppNameCaptured Then
25                    Select Case strUpdateID                 ' Data definition in Update.txt
                        Case "APP_NAME"
26
27                            mstrAppName = strUpdateID_Param
28                            Debug.Print "1> mstrAppName = " & strUpdateID_Param
29                        Case "APP_NEW_FILE_VERSION"
30                            mstrAppNewFileVersion = strUpdateID_Param
31                            Debug.Print "2> mstrAppNewFileVersion = " & strUpdateID_Param & vbCrLf & _
                                        "   gstrAppCurrentFileVer = " & gstrAppCurrentFileVer & vbCrLf & _
                                        "   aeUpdateDebug = " & aeUpdateDebug
311                            gstrLoaderUpdateAppFile = mstrAppNewFileVersion
312                            'MsgBox "312 gstrLoaderUpdateAppFile = " & gstrLoaderUpdateAppFile, vbInformation, _
                                    "aeLoaderUpdateTxtClass: blnTheAppLoaderUpdateStatus"
                              ' If new version does not exist locally then update
32                            If FileExists(gstrLocalPath & mstrAppNewFileVersion) Then
33                                If aeUpdateDebug Then
331                                   'MsgBox "331" & gstrLocalPath & gstrDebugFile, vbInformation, "gconTHIS_APP_VERSION = " & gconTHIS_APP_VERSION
34                                    aeDebugIt gstrLocalPath & gstrDebugFile, "mstrAppNewFileVersion = " & mstrAppNewFileVersion
35                                    aeDebugIt gstrLocalPath & gstrDebugFile, "gstrAppCurrentFileVer = " & gstrAppCurrentFileVer
36                                    aeDebugIt gstrLocalPath & gstrDebugFile, "NOT UPDATING: No change to file version name"
37                                    blnTheAppLoaderUpdateStatus = False
38                                    GoTo Exit_blnTheAppLoaderUpdateStatus
381                                Else
382                                   blnTheAppLoaderUpdateStatus = False
383                                   GoTo Exit_blnTheAppLoaderUpdateStatus
39                                End If
391                           Else
40                            End If
41                        Case "APP_UPDATE_USER"
411                           'MsgBox "411 APP_UPDATE_USER strUpdateID_Param = " & "'" & strUpdateID_Param & "'", vbInformation, _
                                    "aeLoaderUpdateTxtClass: blnTheAppLoaderUpdateStatus"
412                           If strUpdateID_Param = "" Then
413                             blnTheAppLoaderUpdateStatus = False
414                             GoTo Exit_blnTheAppLoaderUpdateStatus
415                           Else
42                                  mastrUpdateUsers = GetUserList(UCase(strUpdateID_Param))
43                                  Debug.Print "3> mastrUpdateUsers = "
44                                  Dim i As Integer
45                                  For i = 0 To UBound(mastrUpdateUsers)
46                                      Debug.Print mastrUpdateUsers(i)
47                                  Next i
                                    'IF NO DEFINED UPDATE USER THEN DO NOT UPDATE
48                                  If mastrUpdateUsers(0) = "NONE" Then
49                                      blnTheAppLoaderUpdateStatus = False
50                                      GoTo Exit_blnTheAppLoaderUpdateStatus
51                                  End If
                                    'IF NOT ALL THEN TEST FOR LEGITIMATE GROUP NAME THEN TEST FOR LEGITIMATE USER NAME
52                                  If mastrUpdateUsers(0) <> "ALL" Then
53                                      mstrTheCurrentUser = UCase(gstrTheCurrentUser)
54                                      Debug.Print "4> mstrTheCurrentUser = " & mstrTheCurrentUser
                                        ' CHECK ALL USERS
55                                      Dim j As Integer
56                                      For j = 0 To UBound(mastrUpdateUsers)
57                                          If mastrUpdateUsers(j) <> mstrTheCurrentUser Then
58                                              Debug.Print mastrUpdateUsers(j)
59                                              blnTheAppLoaderUpdateStatus = False
60                                          Else
61                                              Debug.Print mastrUpdateUsers(j) & " WILL BE UPDATED"
62                                              blnTheAppLoaderUpdateStatus = True
63                                              Exit For
64                                          End If
65                                      Next j
66                                  End If
                              End If
67                        Case "APP_UPDATE_END"
68                            Debug.Print "5> " & "APP_UPDATE_END"
69                            fAppNameCaptured = False    ' Drop out of loop
70                        Case Else
71                    End Select
72                End If
73            End If
74            If Not Comment(strFileData) And fAppNameCaptured Then        ' Create the update text string
75                mstrUpdateText = mstrUpdateText & strFileData & vbCrLf
751               Debug.Print "Loop" & intLoop, mstrUpdateText
76            End If
761           intLoop = intLoop + 1
77        Loop
78    End If
      '
79    gstrUpdateText = mstrUpdateText             ' Store value in frmLoaderUpdateNotes
80    gstrAppNewFileVersion = mstrAppNewFileVersion
      '
81    If blnTheAppLoaderUpdateStatus Then
82        If aeUpdateDebug Then
             'MsgBox "82 aeLoaderUpdateTxtClass: blnTheAppLoaderUpdateStatus", vbInformation
83           aeDebugIt gstrLocalPath & gstrDebugFile, vbCrLf & _
                "gstrTheApp = " & gstrTheApp & vbCrLf & _
                "gstrServerPath = " & gstrServerPath & vbCrLf & _
                "gstrLocalPath = " & gstrLocalPath & vbCrLf & _
                "gstrUpdateInfoFile = " & gstrUpdateInfoFile & vbCrLf & _
                "gstrAppCurrentFileVer = " & gstrAppCurrentFileVer & vbCrLf & _
                "gstrUpdateAppFile = " & gstrUpdateAppFile & vbCrLf & _
                "gstrLoaderUpdateAppFile = " & gstrLoaderUpdateAppFile & vbCrLf & _
                "gstrDebugFile = " & gstrDebugFile & vbCrLf & _
                "gfUpdateDebug = " & gfUpdateDebug & vbCrLf & _
                "_______________________________________________" & vbCrLf & _
                "mstrUpdateText:" & vbCrLf & _
                mstrUpdateText & vbCrLf & _
                "_______________________________________________"
84            aeOpenFrmLoaderUpdateNotes
85            Close #intFileNum                   ' Close data file.
86        Else
87            aeOpenFrmLoaderUpdateNotes
88        End If
89        DoCmd.Quit
90    End If
    
Exit_blnTheAppLoaderUpdateStatus:
91    Exit Function

Err_blnTheAppLoaderUpdateStatus:
92    MsgBox "aeLoaderUpdateTxtClass: blnTheAppLoaderUpdateStatus Erl:" & Erl & " Error # " & Err.Number & ": " & Err.Description
93    MsgBox "gstrTheApp = " & gstrTheApp & vbCrLf & _
            "gstrServerPath = " & gstrServerPath & vbCrLf & _
            "gstrLocalPath = " & gstrLocalPath & vbCrLf & _
            "gstrUpdateInfoFile = " & gstrUpdateInfoFile & vbCrLf & _
            "gstrAppCurrentFileVer = " & gstrAppCurrentFileVer & vbCrLf & _
            "gstrUpdateAppFile = " & gstrUpdateAppFile & vbCrLf & _
            "gstrLoaderUpdateAppFile = " & gstrLoaderUpdateAppFile & vbCrLf & _
            "gstrDebugFile = " & gstrDebugFile & vbCrLf & _
            "gfUpdateDebug = " & gfUpdateDebug & vbCrLf
94    Resume Exit_blnTheAppLoaderUpdateStatus

End Function