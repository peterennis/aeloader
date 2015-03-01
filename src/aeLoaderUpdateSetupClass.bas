Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

' 05/09/2007 - v4.2.3 - Get rid of intDebug, use Erl
' 05/18/2007 - v4.2.4 - Introduce gstrLocalLibPath
'

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
    (ByVal lpBuffer As String, nSize As Long) As Long

Private Declare Function GetComputerName Lib "kernel32.dll" Alias "GetComputerNameA" _
    (ByVal lpBuffer As String, nSize As Long) As Long

Public Enum aeConnectType
    aeWindowsNetworkLogin = 1
    aeMicrosoftAccessLogin = 2
    aeSQLserverLogin = 3
End Enum

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

On Error GoTo Err_aeGetParameter

    Dim dbs As DAO.Database
    Dim rst As DAO.Recordset
    Dim sql As String

1:    sql = "SELECT aeParameters_Table." & TheVarName & " " & _
            "FROM aeParameters_Table " & _
            "WHERE aeParameters_Table.gstrAppName='" & TheApp & "' " & _
            "WITH OWNERACCESS OPTION;"
    'Debug.Print sql

2:    Set dbs = CodeDb()
3:    'Retrieve the data from the database.
4:    Set rst = dbs.OpenRecordset(sql)

    'Debug.Print rst.Fields(0)
5:    aeGetParameter = rst.Fields(0)
6:    Set dbs = Nothing
7:    Set rst = Nothing

Exit_aeGetParameter:
    Exit Function

Err_aeGetParameter:
    MsgBox Err.Description, vbCritical, "DbLib: aeUpdateSetupClass aeGetParameter Error " & Err & _
                "; Erl = " & Erl
    Resume Exit_aeGetParameter
    
End Function

Public Function aeUpdateSetup(ByVal strAppName As String, _
                ByVal strAppCurrentVer As String, _
                ByVal intLoginType As aeConnectType) As Boolean

On Error GoTo Err_aeUpdateSetup

1:    gstrDbLibVersion = aeGetParameter(strAppName, "gstrDbLibVersion")
2:    gstrDbLibName = aeGetParameter(strAppName, "gstrDbLibName")
3:    gstrTheCurrentUser = GetTheCurrentUser(intLoginType)
4:    gstrComputerName = aedblib_GetComputerName()
5:    gfUpdateDebug = aeUpdateDebug
6:    gstrAppCurrentVer = strAppCurrentVer
7:    gstrAppName = strAppName

'    gstrAppName As String        ' String stores the application name.
'    gstrServerPath As String     ' String stores the server path for linked files.
'    gstrLocalPath As String      ' String stores the local application path.
'    gstrLocalLibPath As String   ' String stores the local library path.
'    gstrUpdateInfoFile As String ' String stores the name of the update information file.
'    gstrAppCurrentVer As String  ' String stores the application current version e.g. 4.0.1
'    gstrUpdateAppFile As String  ' String stores the name of the update application file.
'    gstrDebugFile As String      ' String stores the name of the debug file.
'    gfUpdateDebug As Boolean     ' Boolean to turn on debug output.

8:    gstrServerPath = aeGetParameter(strAppName, "gstrServerPath")
9:    gstrLocalPath = aeGetParameter(strAppName, "gstrLocalPath")
10:    gstrLocalLibPath = aeGetParameter(strAppName, "gstrLocalLibPath")
11:    gstrUpdateInfoFile = aeGetParameter(strAppName, "gstrUpdateInfoFile")
12:    gstrUpdateAppFile = aeGetParameter(strAppName, "gstrUpdateAppFile")
13:    gstrDebugFile = aeGetParameter(strAppName, "gstrDebugFile")
14:    gstrUpdateMdb = gstrServerPath & "aeUpdates.mdb"
15:    aeUpdateSetup = True

Exit_aeUpdateSetup:
    Exit Function

Err_aeUpdateSetup:
    MsgBox Err.Description, vbCritical, "DbLib: aeUpdateSetupClass aeUpdateSetup Error " & Err & _
                "; Erl = " & Erl
    Resume Exit_aeUpdateSetup
    
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