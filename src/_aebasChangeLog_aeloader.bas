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
Public Const gstrDATE_aeloader As String = "February 28, 2015"
Public Const gstrVERSION_aeloader As String = "0.0.3"
Public Const gstrPROJECT_aeloader As String = "aeloader"
Public Const gblnTEST_aeloader As Boolean = False
'

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

'=============================================================================================================================
' Tasks:
' %015 -
' %014 -
' %013 -
' %012 -
' %011 -
' %010 -
' Issues:
' #010 -
' #009 -
' #008 -
' #007 -
' #006 -
' #005 -
' #004 -
' #003 -
' #002 -
' #001 -
'=============================================================================================================================
'
'
'20150228 v003 -
    ' FIXED - %009 - Import external data from "aeloader.mdb.v114" then export with aegit, delete from 2do folder
'20150228 v002 -
    ' FIXED - #007 - Import external data from "aeloader.mdb.v112" then export with aegit, delete from 2do folder
    ' FIXED - %008 - Import external data from "aeloader.mdb.v113" then export with aegit, delete from 2do folder
'20150225 v002 -
    ' FIXED - %003 - Import external data from "aeloader.mdb.v106" then export with aegit, delete from 2do folder
    ' FIXED - %004 - Import external data from "aeloader.mdb.v107" then export with aegit, delete from 2do folder
    ' FIXED - %005 - Import external data from "aeloader.mdb.v108" then export with aegit, delete from 2do folder
    ' FIXED - #006 - Import external data from "aeloader.mdb.v111" then export with aegit, delete from 2do folder
'20150223 v001 - First version commit
    ' FIXED - %001 - Import external data from "aeloader not secure.mdb" then export with aegit, delete from 2do folder
    ' FIXED - %002 - Import external data from "aeloader.mdb.v105" then export with aegit, delete from 2do folder