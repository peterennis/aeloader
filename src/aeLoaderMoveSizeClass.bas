Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

' Ref: http://support.microsoft.com/?kbid=210118

Private Declare PtrSafe Function apiGetActiveWindow Lib "user32" Alias "GetActiveWindow" () As Long
Private Declare PtrSafe Function apiGetParent Lib "user32" Alias "GetParent" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function apiShowWindow Lib "user32" Alias "ShowWindow" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Function apiIsIconic Lib "user32" Alias "IsIconic" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function apiIsZoomed Lib "user32" Alias "IsZoomed" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function apiMoveWindow Lib "user32" Alias "MoveWindow" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare PtrSafe Function WM_apiGetDeviceCaps Lib "gdi32" Alias "GetDeviceCaps" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function WM_apiGetDesktopWindow Lib "user32" Alias "GetDesktopWindow" () As Long
Private Declare PtrSafe Function WM_apiGetDC Lib "user32" Alias "GetDC" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function WM_apiReleaseDC Lib "user32" Alias "ReleaseDC" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare PtrSafe Function WM_apiGetSystemMetrics Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long

Private Const SW_SHOWNORMAL = 1
Private Const SW_SHOWMINIMIZED = 2
Private Const SW_MAXIMIZE = 3

Private Const WM_HORZRES = 8
Private Const WM_VERTRES = 10

Private Const LOGPIXELSX = 88
Private Const LOGPIXELSY = 90

Public Property Let aeMoveSizeCenter(bln As Boolean)
' Move app window center screen and smallest size
    MoveSizeCenter
End Property

Private Sub MoveSizeCenter()

    Dim intScreenCenterHorizontalPix As Integer
    Dim intScreenCenterVerticalPix As Integer

'    Debug.Print "GetAccesshWnd=" & GetAccesshWnd
'    AccessMinimize
'    AccessRestore
'    AccessMaximize
'    Debug.Print "IsAccessMaximized=" & IsAccessMaximized
'    Debug.Print "IsAccessRestored=" & IsAccessRestored
'    Debug.Print "IsAccessMinimized=" & IsAccessMinimized
    intScreenCenterHorizontalPix = aefGetScreenResolution("H") \ 2
    intScreenCenterVerticalPix = aefGetScreenResolution("V") \ 2
    Debug.Print "intScreenCenterHorizontalPix=" & intScreenCenterHorizontalPix
    Debug.Print "intScreenCenterVerticalPix=" & intScreenCenterVerticalPix
    AccessMoveSize intScreenCenterHorizontalPix, intScreenCenterVerticalPix, 0, 0

End Sub

Private Sub ScreenResolutionTest()

    Dim lngScreenCenterHorizontalTwips As Long
    Dim lngScreenCenterVerticalTwips As Long
    lngScreenCenterHorizontalTwips = (aefPixelsToTwips(aefGetScreenResolution("H"), "H")) \ 2
    lngScreenCenterVerticalTwips = (aefPixelsToTwips(aefGetScreenResolution("V"), "V")) \ 2

    'Debug.Print "lngScreenCenterHorizontalTwips=" & lngScreenCenterHorizontalTwips
    'Debug.Print "lngScreenCenterVerticalTwips=" & lngScreenCenterVerticalTwips

End Sub

Private Function GetAccesshWnd()
    Dim hWnd As Long
    Dim hWndAccess As Long

    ' Get the handle to the currently active window.
    hWnd = apiGetActiveWindow()
    hWndAccess = hWnd

    ' Find the top window (which has no parent window).
    While hWnd <> 0
        hWndAccess = hWnd
        hWnd = apiGetParent(hWnd)
    Wend

    GetAccesshWnd = hWndAccess

End Function

Private Function AccessMinimize()
    AccessMinimize = apiShowWindow(GetAccesshWnd(), SW_SHOWMINIMIZED)
End Function

Private Function AccessMaximize()
    AccessMaximize = apiShowWindow(GetAccesshWnd(), SW_MAXIMIZE)
End Function

Private Function AccessRestore()
    AccessRestore = apiShowWindow(GetAccesshWnd(), SW_SHOWNORMAL)
End Function

Private Function IsAccessMaximized()
    If apiIsZoomed(GetAccesshWnd()) = 0 Then
        IsAccessMaximized = False
    Else
        IsAccessMaximized = True
    End If
End Function

Private Function IsAccessMinimized()
    If apiIsIconic(GetAccesshWnd()) = 0 Then
        IsAccessMinimized = False
    Else
        IsAccessMinimized = True
    End If
End Function

Private Function IsAccessRestored()
    If IsAccessMaximized() = False And IsAccessMinimized() = False Then
        IsAccessRestored = True
    Else
        IsAccessRestored = False
    End If
End Function

Private Sub AccessMoveSize(iX As Integer, iY As Integer, iWidth As Integer, iHeight As Integer)
    apiMoveWindow GetAccesshWnd(), iX, iY, iWidth, iHeight, True
End Sub

Private Function aefGetScreenResolution(Optional strHorV As Variant) As String
'Ref: http://www.peterssoftware.com/c_scrres.htm

    ' Return the display height and width
    Dim DisplayHeight As Integer
    Dim DisplayWidth As Integer
    Dim hDesktopWnd As Long
    Dim hDCcaps As Long
    Dim iRtn As Integer

    ' Make API calls to get desktop settings
    hDesktopWnd = WM_apiGetDesktopWindow()  ' Get handle to desktop
    hDCcaps = WM_apiGetDC(hDesktopWnd)      ' Get Display Context
    DisplayHeight = WM_apiGetDeviceCaps(hDCcaps, WM_VERTRES)
    DisplayWidth = WM_apiGetDeviceCaps(hDCcaps, WM_HORZRES)
    iRtn = WM_apiReleaseDC(hDesktopWnd, hDCcaps)

    If IsMissing(strHorV) Then
        aefGetScreenResolution = DisplayWidth & "x" & DisplayHeight
    ElseIf UCase(Left(strHorV, 1)) = "H" Then
        aefGetScreenResolution = DisplayWidth
    ElseIf UCase(Left(strHorV, 1)) = "V" Then
        aefGetScreenResolution = DisplayHeight
    Else
        MsgBox "aeGetScreenResolution(""H"") returns the horizontal resolution as a string, " & vbCrLf & _
            "aeGetScreenResolution(""V"") returns the vertical resolution as a string, " & vbCrLf & _
            "aeGetScreenResolution() returns the resolution as a string DisplayWidth x DisplayHeight.", _
            vbCritical, "aeGetScreenResolution"
    End If

End Function

Private Function aefTwipsToPixels(lngTwips As Long, strHorV As Variant) As Long
' Ref: http://www.applecore99.com/api/api012.asp
' Adapted from Q94927 in the Microsoft Knowledge Base
'
'   Function to convert Twips to pixels for the current screen resolution
'   Accepts:
'       lngTwips - the number of twips to be converted
'       lngDirection - direction (x or y - use either DIRECTION_VERTICAL or DIRECTION_HORIZONTAL)
'   Returns:
'       the number of pixels corresponding to the given twips

    On Error GoTo PROC_ERR

    Dim lngDeviceHandle As Long
    Dim lngPixelsPerInch As Long
    lngDeviceHandle = WM_apiGetDC(0)
    If UCase(Left(strHorV, 1)) = "H" Then       ' Horizontal
        lngPixelsPerInch = WM_apiGetDeviceCaps(lngDeviceHandle, LOGPIXELSX)
    ElseIf UCase(Left(strHorV, 1)) = "V" Then   ' Vertical
        lngPixelsPerInch = WM_apiGetDeviceCaps(lngDeviceHandle, LOGPIXELSY)
    Else
        MsgBox "Use ""H"" (horizontal) or ""V"" (vertical) for direction. ", _
            vbCritical, "fTwipsToPixels"
        Exit Function
    End If
    lngDeviceHandle = WM_apiReleaseDC(0, lngDeviceHandle)
    aefTwipsToPixels = lngTwips / 1440 * lngPixelsPerInch

PROC_EXIT:
    On Error Resume Next
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aefTwipsToPixels of Class aeLoaderMoveSizeClass", _
                vbOKOnly + vbCritical, gconTHIS_APP_NAME
    Resume PROC_EXIT

End Function

Private Function aefPixelsToTwips(lngPixels As Long, strHorV As Variant) As Long
'   Function to convert pixels to twips for the current screen resolution
'   Accepts:
'       lngPixels - the number of pixels to be converted
'       lngDirection - direction (x or y - use either DIRECTION_VERTICAL or DIRECTION_HORIZONTAL)
'   Returns:
'       the number of twips corresponding to the given pixels

    On Error GoTo PROC_ERR
    
    Dim lngDeviceHandle As Long
    Dim lngPixelsPerInch As Long
    lngDeviceHandle = WM_apiGetDC(0)
    If UCase(Left(strHorV, 1)) = "H" Then       ' Horizontal
        lngPixelsPerInch = WM_apiGetDeviceCaps(lngDeviceHandle, LOGPIXELSX)
    ElseIf UCase(Left(strHorV, 1)) = "V" Then   ' Vertical
        lngPixelsPerInch = WM_apiGetDeviceCaps(lngDeviceHandle, LOGPIXELSY)
    Else
        MsgBox "Use ""H"" (horizontal) or ""V"" (vertical) for direction. ", _
            vbCritical, "fTwipsToPixels"
        Exit Function
    End If
    lngDeviceHandle = WM_apiReleaseDC(0, lngDeviceHandle)
    aefPixelsToTwips = lngPixels * 1440 / lngPixelsPerInch

PROC_EXIT:
    On Error Resume Next
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aefPixelsToTwips of Class aeLoaderMoveSizeClass", _
                vbOKOnly + vbCritical, gconTHIS_APP_NAME
    Resume PROC_EXIT

End Function