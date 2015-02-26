Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

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
    Dim intDebug As Integer

1    sql = "SELECT aeParameters_Table." & TheVarName & " " & _
            "FROM aeParameters_Table " & _
            "WHERE aeParameters_Table.gstrAppName='" & TheApp & "' " & _
            "WITH OWNERACCESS OPTION;"
    'Debug.Print sql
    
2    Set dbs = CurrentDb()
    'Retrieve the data from the database.
3    Set rst = dbs.OpenRecordset(sql)

    'Debug.Print rst.Fields(0)
4    aeGetParameter = rst.Fields(0)
5    Set dbs = Nothing
6    Set rst = Nothing

Exit_aeGetParameter:
7    Exit Function

Err_aeGetParameter:
8    MsgBox "Erl" & Erl & " Error# " & Err & " " & Err.Description, _
            vbCritical, "aeLoader: aeUpdateSetupClass aeGetParameter"
9    Resume Exit_aeGetParameter
    
End Function

Public Function GetAppCurrentFileVer() As String

    If FileExists(gstrLocalPath & gstrTheApp) Then
        GetAppCurrentFileVer = gstrTheApp
    End If
    
End Function

Public Function aeUpdateSetup(ByVal strAppName As String, _
                ByVal strAppCurrentVer As String, _
                ByVal intLoginType As aeConnectType) As Boolean

On Error GoTo Err_aeUpdateSetup

'    gstrTheApp As String         ' String stores the application name.
'    gstrServerPath As String     ' String stores the server path for linked files.
'    gstrLocalPath As String      ' String stores the local application path.
'    gstrUpdateInfoFile As String ' String stores the name of the update information file.
'    gstrAppCurrentFileVer As String    ' String stores the application current filename version
'    gstrLoaderUpdateAppFile As String  ' String stores the name of the loader update application file.
'    gstrDebugFile As String      ' String stores the name of the debug file.
'    gfUpdateDebug As Boolean     ' Boolean to turn on debug output.

    Dim intDebug As Integer

1    gstrTheAppWindowName = Nz(DLookup("gconAPP_WINDOW_NAME", "tblAppSetup", _
                            "AppID=" & gintApp))
2    gstrPassThrough = Nz(DLookup("gconPASS_THROUGH", "tblAppSetup", _
                            "AppID=" & gintApp))
3    gstrLocalPath = Nz(DLookup("gconLOCAL_PATH", "tblAppSetup", _
                            "AppID=" & gintApp))
4    '
5    gstrTheAppExtension = Nz(DLookup("gconAPP_EXTENSION", "tblAppSetup", _
                            "AppID=" & gintApp))
6    gstrTheWorkgroup = Nz(DLookup("gconSERVER_PATH", "tblAppSetup", _
                            "AppID=" & gintApp)) & _
                            Nz(DLookup("gconTHE_WORKGROUP", "tblAppSetup", _
                            "AppID=" & gintApp))
7    gstrTheApp = Nz(DLookup("gconAPP_NAME", "tblAppSetup", _
                            "AppID=" & gintApp))
701  gstrTheAppSeparatorChar = Nz(DLookup("gconAPP_SEPARATOR", "tblAppSetup", _
                            "AppID=" & gintApp))
8    gstrLogonMdb = Nz(DLookup("gconLOGON_MDB", "tblAppSetup", _
                            "AppID=" & gintApp))
9    gstrPasswordMdb = Nz(DLookup("gconPASSWORD_MDB", "tblAppSetup", _
                            "AppID=" & gintApp))
10 'MsgBox "gstrTheAppWindowName = " & gstrTheAppWindowName & vbCrLf & _
            "gstrPassThrough = " & gstrPassThrough & vbCrLf & _
            "gstrLocalPath = " & gstrLocalPath & vbCrLf & _
            "gstrTheAppExtension = " & gstrTheAppExtension & vbCrLf & _
            "gstrTheWorkgroup = " & gstrTheWorkgroup & vbCrLf & _
            "gstrTheApp = " & gstrTheApp & vbCrLf & _
            "gstrTheAppSeparatorChar = " & gstrTheAppSeparatorChar & vbCrLf & _
            "gstrLogonMdb = " & gstrLogonMdb & vbCrLf & _
            "gstrPasswordMdb = " & gstrPasswordMdb

11    gstrTheCurrentUser = GetTheCurrentUser(intLoginType)
12    gstrComputerName = aeLoaderGetComputerName()
13    gfUpdateDebug = aeUpdateDebug
131   gstrAppCurrentFileVer = GetAppCurrentFileVer
14    gstrServerPath = Nz(DLookup("gconSERVER_PATH", "tblAppSetup", _
                            "AppID=" & gintApp))
15    gstrUpdateInfoFile = Nz(DLookup("gconUPDATE_INFO_FILE", "tblAppSetup", _
                            "AppID=" & gintApp))
16    gstrUpdateAppFile = Nz(DLookup("gconUPDATE_APP_FILE", "tblAppSetup", _
                            "AppID=" & gintApp))
17    gstrDebugFile = "aeLoaderDebugFile.txt"
'x    gstrUpdateMdb = gstrServerPath & "aeUpdates.mdb"
    
18 'MsgBox "gstrTheCurrentUser = " & gstrTheCurrentUser & vbCrLf & _
            "gstrComputerName = " & gstrComputerName & vbCrLf & _
            "gfUpdateDebug = " & gfUpdateDebug & vbCrLf & _
            "gstrAppCurrentFileVer = " & gstrAppCurrentFileVer & vbCrLf & _
            "gstrServerPath = " & gstrServerPath & vbCrLf & _
            "gstrLocalPath = " & gstrLocalPath & vbCrLf & _
            "gstrUpdateInfoFile = " & gstrUpdateInfoFile & vbCrLf & _
            "gstrUpdateAppFile = " & gstrUpdateAppFile & vbCrLf & _
            "gstrLoaderUpdateAppFile = " & gstrLoaderUpdateAppFile & vbCrLf & _
            "gstrDebugFile = " & gstrDebugFile

19    aeUpdateSetup = True
    
Exit_aeUpdateSetup:
20    Exit Function

Err_aeUpdateSetup:
21    MsgBox Err.Description, vbCritical, "aeLoader: aeUpdateSetupClass aeUpdateSetup Erl:" & Erl & " Error# " & Err & " " & Err.Description
22    Resume Exit_aeUpdateSetup

End Function

Private Function aeLoaderGetComputerName() As Variant
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
    aeLoaderGetComputerName = Left(strComputerName, InStr(1, strComputerName, _
                                Chr(0)) - 1)
    
End Function

Private Function aeLoaderGetUserName() As Variant
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
    aeLoaderGetUserName = Left(strUserName, InStr(1, strUserName, _
                                Chr(0)) - 1)
    
End Function

Private Static Property Get GetTheCurrentUser( _
        ByVal intTypeOfUserConnection As aeConnectType) As String
    
    Select Case intTypeOfUserConnection
        Case aeConnectType.aeWindowsNetworkLogin:   ' 1
            ' Windows Network Login
            GetTheCurrentUser = aeLoaderGetUserName
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