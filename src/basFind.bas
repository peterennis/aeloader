Option Compare Database
Option Explicit

' Ref   : Modified from http://support.microsoft.com/kb/q185476/
' What  : Get and array of filenames
' Who   : Peter F. Ennis
' When  : 06/06/2007

Private mblnFoundTheFile As Boolean

Private Function aeFindFiles(ByVal path As Variant, ByVal SearchStr As String, _
            ByRef FileCount As Integer, ByRef DirCount As Integer, _
            ByRef aFileNames() As String)
            
' e.g.: aeFindFiles("C:\Windows\","*.*",FileCountResult,DirCountResult)
      
' NOTE: Using "C:\My Documents\" does not work on my Win2K.
'           "The system cannot find the path secified" error.
'           No simple answer, so leave it.

      Dim FileName As String   ' Walking filename variable.
      Dim DirName As String    ' SubDirectory Name.
      Dim dirNames() As String ' Buffer for directory name entries.
      Dim nDir As Integer      ' Number of directories in this path.
      Dim i As Integer         ' For-loop counter.
      Dim SizeOfArray As Integer
      
On Error GoTo Err_aeFindFiles
      
1:      If Right(path, 1) <> "\" Then path = path & "\"
      ' Search for subdirectories.
2:      nDir = 0
3:      ReDim dirNames(nDir)
4:      DirName = Dir(path, vbDirectory Or vbHidden Or vbArchive Or vbReadOnly _
            Or vbSystem)  ' Even if hidden, and so on.
5:      Do While Len(DirName) > 0
         ' Ignore the current and encompassing directories.
6:         If (DirName <> ".") And (DirName <> "..") Then
            ' Check for directory with bitwise comparison.
7:            If GetAttr(path & DirName) And vbDirectory Then
8:               dirNames(nDir) = DirName
9:               DirCount = DirCount + 1
10:               nDir = nDir + 1
11:               ReDim Preserve dirNames(nDir)
               'List2.AddItem path & DirName ' Uncomment to list directories.
12:               Debug.Print "Dir=" & path & DirName
13:            End If
sysFileERRCont:
14:         End If
15:         DirName = Dir()  ' Get next subdirectory.
16:      Loop

      ' Search through this directory and sum file sizes.
17:      FileName = Dir(path & SearchStr, vbNormal Or vbHidden Or vbSystem _
            Or vbReadOnly Or vbArchive)
18:      While Len(FileName) <> 0
19:          aeFindFiles = aeFindFiles + FileLen(path & FileName)
'         ' Load List box
'         List2.AddItem path & FileName & vbTab & _
            FileDateTime(path & FileName)   ' Include Modified Date
''#PFE#          Debug.Print path & FileName & vbTab & _
''              FileDateTime(path & FileName)   ' Include Modified Date
20:          ReDim Preserve aFileNames(FileCount)
21:          aFileNames(FileCount) = FileName
''#PFE#          Debug.Print "A:", FileCount, aFileNames(FileCount)
22:          FileCount = FileCount + 1
23:          FileName = Dir()  ' Get next file.
24:      Wend

      ' If there are sub-directories..
25:      If nDir > 0 Then
         ' Recursively walk into them
26:         For i = 0 To nDir - 1
27:           aeFindFiles = aeFindFiles + aeFindFiles(path & dirNames(i) & "\", _
                 SearchStr, FileCount, DirCount, aFileNames())
28:         Next i
29:      End If

30:     If FileCount = 0 Then
31:         'MsgBox "No files found!"
32:         mblnFoundTheFile = False
33:     Else
34:         mblnFoundTheFile = True
35:     End If
    
Exit_aeFindFiles:
      Exit Function

Err_aeFindFiles:
      If Right(DirName, 4) = ".sys" Then
        Resume sysFileERRCont ' Known issue with pagefile.sys
      Else
        MsgBox "Erl=" & Erl & " - " & Err.Description, vbCritical, _
             "Err_aeFindFiles Unexpected Error Err=" & Err.Number
        Resume Exit_aeFindFiles
      End If

End Function

Private Sub ListAllTheFiles(ByRef strTheNewLibFile As String)
      
    Dim SearchPath As String
    Dim FindStr As String
    Dim FileSize As Long
    Dim NumFiles As Integer
    Dim NumDirs As Integer
    Dim aFNameResults() As String
    Dim i As Integer
    Dim strMessage As String

On Error GoTo Err_ListAllTheFiles

0:    mblnFoundTheFile = False
1:    DoCmd.Hourglass True
'      List2.Clear
2:    SearchPath = "C:\DSFRC\INTAKE\"
3:    FindStr = "adaeptdblib.mda.v*"
4:    FileSize = aeFindFiles(SearchPath, FindStr, NumFiles, NumDirs, aFNameResults())
5:    Debug.Print NumFiles & " Files found in " & NumDirs + 1 & _
                                " Directories"
6:    Debug.Print "Size of files found under " & SearchPath & " = " & _
            Format(FileSize, "#,###,###,##0") & " Bytes"
7:    DoCmd.Hourglass False

8:    If mblnFoundTheFile Then
9:      If UBound(aFNameResults()) > 1 Then
10:        For i = 0 To UBound(aFNameResults())
11:            Debug.Print i, aFNameResults(i)
12:            strMessage = strMessage & aFNameResults(i) & vbCrLf
13:        Next i
14:        MsgBox strMessage & vbCrLf & "Too many library files!" & vbCrLf & _
                "Please contact the system administrator.", vbCritical, "ListAllFiles"
15:      Else
          'MsgBox "The new library file is " & aFNameResults(i), vbInformation, "ListAllFiles"
16:        strTheNewLibFile = aFNameResults(i)
17:      End If
18:    End If

Exit_ListAllTheFiles:
    Exit Sub

Err_ListAllTheFiles:
    MsgBox "Erl=" & Erl & " " & Err.Description, vbCritical, "Err_ListAllTheFiles Err=" & Err
    Resume Exit_ListAllTheFiles
      
End Sub

Public Sub InstallNewLibrary()

    Dim strLibFile As String
    
On Error GoTo Err_InstallNewLibrary

1:    ListAllTheFiles strLibFile
2:    Debug.Print "The new lib file is: " & strLibFile
3:    'MsgBox "The new lib file is: " & strLibFile
'
4:    If mblnFoundTheFile Then

        ' Make backup copy of old library file
5:     If FileExists(gstrLocalLibPath & "adaeptdblib.mda.OLD") Then
6:        Kill gstrLocalLibPath & "adaeptdblib.mda.OLD"
7:     End If
8:    'MsgBox "gstrLocalLibPath & 'adaeptdblib.mda'=" & gstrLocalLibPath & "adaeptdblib.mda"
9:    'MsgBox "gstrLocalLibPath & 'adaeptdblib.mda.OLD'=" & gstrLocalLibPath & "adaeptdblib.mda.OLD"
10:    If FileExists(gstrLocalLibPath & "adaeptdblib.mda") Then
11:        Name gstrLocalLibPath & "adaeptdblib.mda" _
                As gstrLocalLibPath & "adaeptdblib.mda.OLD"
12:     End If
13:      If FileExists(gstrLocalLibPath & strLibFile) Then
14:          Name gstrLocalLibPath & strLibFile _
                  As gstrLocalLibPath & "adaeptdblib.mda"
15:      End If
16:    End If

Exit_InstallNewLibrary:
    Exit Sub

Err_InstallNewLibrary:
    MsgBox "Erl=" & Erl & " " & Err.Description, vbCritical, "Err_InstallNewLibrary Err=" & Err
    Resume Exit_InstallNewLibrary

End Sub