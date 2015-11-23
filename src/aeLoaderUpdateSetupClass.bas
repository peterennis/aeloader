Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

Private Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare PtrSafe Function GetComputerName Lib "kernel32.dll" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Enum aeConnectType
    aeWindowsNetworkLogin = 1
    aeMicrosoftAccessLogin = 2
    aeSQLserverLogin = 3
End Enum

Private Sub Class_Initialize()
    Debug.Print "aeLoaderUpdateSetupClass: Class_Initialize"
End Sub

Private Sub Class_Terminate()
    Debug.Print "aeLoaderUpdateSetupClass: Class_Terminate"
End Sub

Public Property Let aeUpdateDebug(bln As Boolean)
' Allow Debug to be turned on outside of the class
    gfUpdateDebug = bln
End Property

Public Property Get aeUpdateDebug() As Boolean
    aeUpdateDebug = gfUpdateDebug
End Property

Private Function aeGetParameter(ByVal TheApp As String, _
            ByVal TheVarName As String) As String
' Ref: http://support.microsoft.com/default.aspx?scid=kb;en-us;Q149254

    Debug.Print "aeGetParameter"

    On Error GoTo PROC_ERR

    Dim dbs As DAO.Database
    Dim rst As DAO.Recordset
    Dim strSQL As String

    strSQL = "SELECT aeLoaderParameters_Table." & TheVarName & " " & _
            "FROM aeLoaderParameters_Table " & _
            "WHERE aeLoaderParameters_Table.gstrAppName='" & TheApp & "' " & _
            "WITH OWNERACCESS OPTION;"
    'MsgBox strSQL & vbCrLf & _
            "TheVarName=" & TheVarName & vbCrLf & _
            "TheApp=" & TheApp, vbInformation, "Here"
    Debug.Print , "strSQL = " & strSQL

    Set dbs = CodeDb()
    ' Retrieve the data from the database
    Set rst = dbs.OpenRecordset(strSQL)

    'Debug.Print rst.Fields(0)
    aeGetParameter = rst.Fields(0)
    rst.Close
    dbs.Close
    Set rst = Nothing
    Set dbs = Nothing

PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Err=" & Err & " " & Err.Description, vbCritical, _
            "aeLoaderUpdateSetupClass: aeGetParameter"
    Resume PROC_EXIT
    
End Function

Public Function aeUpdateSetup(ByVal strAppName As String, _
                ByVal strAppCurrentVer As String, _
                ByVal intLoginType As aeConnectType) As Boolean

    Debug.Print "aeUpdateSetup"

    On Error GoTo PROC_ERR

    gstrDbLibVersion = aeGetParameter(strAppName, "gstrDbLibVersion")
    Debug.Print , "strAppName=" & strAppName & gstrDbLibVersion
    gstrDbLibName = aeGetParameter(strAppName, "gstrDbLibName")
    Debug.Print , "strAppName=" & strAppName & gstrDbLibName
    gstrTheCurrentUser = GetTheCurrentUser(intLoginType)
    Debug.Print , "gstrTheCurrentUser=" & gstrTheCurrentUser
    gstrComputerName = aedblib_GetComputerName()
    Debug.Print , "gstrComputerName=" & gstrComputerName
    gfUpdateDebug = aeUpdateDebug
    Debug.Print , "gfUpdateDebug=" & gfUpdateDebug
    gstrAppCurrentVer = strAppCurrentVer
    Debug.Print , "gstrAppCurrentVer=" & gstrAppCurrentVer
    gstrAppName = strAppName
    Debug.Print , "gstrAppName=" & gstrAppName

'    gstrAppName As String        ' String stores the application name.
'    gstrServerPath As String     ' String stores the server path for linked files.
'    gstrLocalPath As String      ' String stores the local application path.
'    gstrLocalLibPath As String   ' String stores the local library path.
'    gstrUpdateInfoFile As String ' String stores the name of the update information file.
'    gstrAppCurrentVer As String  ' String stores the application current version e.g. 4.0.1
'    gstrUpdateAppFile As String  ' String stores the name of the update application file.
'    gstrDebugFile As String      ' String stores the name of the debug file.
'    gfUpdateDebug As Boolean     ' Boolean to turn on debug output.

    gstrServerPath = aeGetParameter(strAppName, "gstrServerPath")
    Debug.Print , "gstrServerPath=" & gstrServerPath
    gstrLocalPath = aeGetParameter(strAppName, "gstrLocalPath")
    Debug.Print , "gstrLocalPath=" & gstrLocalPath
    gstrLocalLibPath = aeGetParameter(strAppName, "gstrLocalLibPath")
    Debug.Print , "gstrLocalLibPath=" & gstrLocalLibPath
    gstrUpdateInfoFile = aeGetParameter(strAppName, "gstrUpdateInfoFile")
    Debug.Print , "gstrUpdateInfoFile=" & gstrUpdateInfoFile
    gstrUpdateAppFile = aeGetParameter(strAppName, "gstrUpdateAppFile")
    Debug.Print , "gstrUpdateAppFile=" & gstrUpdateAppFile
    gstrDebugFile = aeGetParameter(strAppName, "gstrDebugFile")
    Debug.Print , "gstrDebugFile=" & gstrDebugFile
    gstrUpdateMdb = gstrServerPath & "aeUpdates.mdb"
    Debug.Print , "gstrUpdateMdb=" & gstrUpdateMdb
    aeUpdateSetup = True

PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Err=" & Err & " " & Err.Description, vbCritical, "aeLoaderUpdateSetupClass aeUpdateSetup"
    Resume PROC_EXIT
    
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

Private Function aedblib_GetUserName() As Variant
' Wrapper function for API GetUserNameA routine

    Dim strUserName As String
    Dim lngLength As Long
    Dim lngResult As Long
    
    ' Set up the buffer
    strUserName = String$(255, 0)
    lngLength = 255
    ' Make the call
    lngResult = GetUserName(strUserName, lngLength)
    ' Cleanup and assign the value
    aedblib_GetUserName = Left(strUserName, InStr(1, strUserName, _
                                Chr(0)) - 1)
    
End Function

Private Static Property Get GetTheCurrentUser( _
        ByVal intTypeOfUserConnection As aeConnectType) As String
    
    Select Case intTypeOfUserConnection
        Case aeConnectType.aeWindowsNetworkLogin:   ' 1
            ' Windows Network Login
            GetTheCurrentUser = aedblib_GetUserName
            gstrNetUserLogin = GetTheCurrentUser
        Case aeConnectType.aeMicrosoftAccessLogin:  ' 2
            ' Microsoft Access Login
            GetTheCurrentUser = CurrentUser()
            gstrMdbUserLogin = GetTheCurrentUser
        Case aeConnectType.aeSQLserverLogin:        ' 3
            ' SQL Server Login
            Debug.Print "Get the SQL Server Login"
            GetTheCurrentUser = "Add a Routine to Get the SQL Server Login"
            gstrSqlUserLogin = GetTheCurrentUser
        Case Else
    End Select
        
End Property