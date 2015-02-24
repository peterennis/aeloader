Option Compare Database
Option Explicit

' Custom Usage:
' Public Const THE_SOURCE_FOLDER = "Z:\The\Source\Folder\src.MYPROJECT\"
' Public Const THE_XML_FOLDER = "Z:\The\Source\Folder\src.MYPROJECT\xml\"
' For custom configuration of the output source folder in aegitClassTest use:
' oDbObjects.SourceFolder = THE_SOURCE_FOLDER
' oDbObjects.XMLfolder = THE_XML_FOLDER
' Run in immediate window:                  ALTERNATIVE_EXPORT
' Show debug output in immediate window:    ALTERNATIVE_EXPORT varDebug:="varDebug"
'                                           ALTERNATIVE_EXPORT 1
'
' Sample constants for settings of "TheProjectName"
Public Const gstrDATE_TheProjectName As String = "January 1, 2000"
Public Const gstrVERSION_TheProjectName As String = "0.0.0"
Public Const gstrPROJECT_TheProjectName As String = "TheProjectName"
Public Const gblnTEST_TheProjectName As Boolean = False

Public Function aeloader_EXPORT(Optional ByVal varDebug As Variant) As Boolean

    Dim THE_SOURCE_FOLDER As String
    Dim THE_XML_FOLDER As String
    Dim THE_XML_DATA_FOLDER As String

    THE_SOURCE_FOLDER = "C:\ae\aeloader\src\"
    THE_XML_FOLDER = "C:\ae\aeloader\src\xml"
    THE_XML_DATA_FOLDER = "C:\ae\aeloader\src\xml"
 
    On Error GoTo PROC_ERR
 
    If Not IsMissing(varDebug) Then
         aegitClassTest varDebug:="varDebug", varSrcFldr:=THE_SOURCE_FOLDER, varXmlFldr:=THE_XML_FOLDER
     Else
         aegitClassTest varSrcFldr:=THE_SOURCE_FOLDER, varXmlFldr:=THE_XML_FOLDER
     End If
 
PROC_EXIT:
     Exit Function
 
PROC_ERR:
     MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeloader_EXPORT"
     Resume Next

End Function