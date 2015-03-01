Option Compare Database
Option Explicit

' Ref: http://support.microsoft.com/default.aspx?scid=kb;EN-US;168829

Public Const GW_CHILD = 5
Public Const GW_HWNDNEXT = 2

Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, _
              ByVal wCmd As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
              (ByVal hWnd As Long, ByVal lpString As String, _
              ByVal cch As Long) As Long
Declare Function GetTopWindow Lib "user32" _
              (ByVal hWnd As Long) As Long
Declare Function GetClassName Lib "user32" Alias "GetClassNameA" _
              (ByVal hWnd As Long, ByVal lpClassName As String, _
              ByVal nMaxCount As Long) As Long