Option Compare Database
Option Explicit

Public Function CurDir(Optional Drive As String)
    CurDir = VBA.CurDir(Drive)
End Function

Public Function Environ(Expression)
    Environ = VBA.Environ(Expression)
End Function

Public Sub TEST_aeLoaderUpdate()

    Dim bln As Boolean

    gintApp = 6

    ' Call aeLoaderUpdateSetupClass
    Dim blnUpdate As Boolean
    Dim cls1 As aeLoaderUpdateSetupClass
    Set cls1 = New aeLoaderUpdateSetupClass

    ' Setup parameters
    cls1.aeUpdateDebug = True
    blnUpdate = cls1.aeUpdateSetup(gconTHIS_APP_NAME, gconTHIS_APP_VERSION, aeWindowsNetworkLogin)

    ' Example with Network Login
    Dim cls2 As aeLoaderUpdateTxtClass
    Set cls2 = New aeLoaderUpdateTxtClass
    cls2.aeUpdateDebug = True
    blnUpdate = cls2.blnTheAppLoaderUpdateStatus()
    Debug.Print "cls2.blnTheAppLoaderUpdateStatus = " & blnUpdate

    ' Shutdown the app if it is already open
    'MsgBox "gstrTheAppWindowName = " & gstrTheAppWindowName, vbInformation, gconTHIS_APP_NAME & ": StartApp"
    ShutDownApplication (gstrTheAppWindowName)

    bln = aeLoaderPassThroughApp(gstrLocalPath, gstrLoaderUpdateAppFile)

End Sub