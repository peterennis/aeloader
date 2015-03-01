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
    ' Move notes from old history in basGlobal
    ' (c) 2000 - 2005 adaept Peter Ennis
    ' 09/08/2004 1.0.1 - Add code to minimize Access on startup.
    ' 02/23/2005 1.0.2 - Office 2003 Compatability, GetAccessVersion()
    ' 02/25/2005 1.0.3 - Use tblAppSetup, gintApp, DLookup for flexibility
    ' 03/03/2005 1.0.4 - Debug operation with Medical database
    ' 03/21/2005 1.0.5 - Add SLCC with Logon and Password fields in startup table.
    ' 07/22/2005 1.0.6 - DoCmd.Restore added when application closes.
    ' 08/18/2005 1.0.7 - Add DSFRC Volunteers.
    ' 08/24/2005 1.0.8 - Add DSFRC Finance.
    ' 08/31/2005 1.0.9 - Add NOPWD group and NopwdUser as pass through for Noho.
    ' 09/15/2005 1.1.0 - Pass through test succeeds for NoHo.
    ' 09/16/2005 1.1.1 - Fix bug where pass through messed up original version control.
    ' 06/09/2006 1.1.2 - Add aeLoaderMoveSizeClass to center and reduce access db window
    '                    to its smallest size on any screen.
    ' 05/10/2007 1.1.3 - Allow update with user:PCname and kill .OLD mda file before rename.
    ' 05/18/2007 1.1.4 - Updates to class modules from adaeptdblib.mda
    '                    Debugging to trap file permissions error when user is not in admin group.
    '                    Need to allow the application to delete and rename files.
    ' 05/21/2007 1.1.5 - Import aeParameters_Table from adaeptdblib.mda to fix bug resulting from
    '                    starting Noho with last update of aeloader.
    ' 05/22/2007 1.1.6 - Replace tblAppSetup values with aeLoaderParameters_Table.
    ' 06/07/2007 1.1.7 - Update library e.g. adaeptdblib.mda.v425
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