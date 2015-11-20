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

    Debug.Print "aeFindFiles"

    Dim FileName As String   ' Walking filename variable.
    Dim DirName As String    ' SubDirectory Name.
    Dim dirNames() As String ' Buffer for directory name entries.
    Dim nDir As Integer      ' Number of directories in this path.
    Dim i As Integer         ' For-loop counter.
    Dim SizeOfArray As Integer

    On Error GoTo PROC_ERR

    If Right(path, 1) <> "\" Then path = path & "\"
    ' Search for subdirectories.
    nDir = 0
    ReDim dirNames(nDir)
    DirName = Dir(path, vbDirectory Or vbHidden Or vbArchive Or vbReadOnly _
                    Or vbSystem)  ' Even if hidden, and so on.

    Do While Len(DirName) > 0
    ' Ignore the current and encompassing directories.
        If (DirName <> ".") And (DirName <> "..") Then
        ' Check for directory with bitwise comparison.
            If GetAttr(path & DirName) And vbDirectory Then
                dirNames(nDir) = DirName
                DirCount = DirCount + 1
                nDir = nDir + 1
                ReDim Preserve dirNames(nDir)
                'List2.AddItem path & DirName ' Uncomment to list directories.
                Debug.Print , "Dir=" & path & DirName
            End If
sysFileERRCont:
        End If
        DirName = Dir()  ' Get next subdirectory
    Loop

    ' Search through this directory and sum file sizes.
    FileName = Dir(path & SearchStr, vbNormal Or vbHidden Or vbSystem _
                    Or vbReadOnly Or vbArchive)

    While Len(FileName) <> 0
        aeFindFiles = aeFindFiles + FileLen(path & FileName)
'         ' Load List box
'         List2.AddItem path & FileName & vbTab & _
      FileDateTime(path & FileName)   ' Include Modified Date
''#PFE#          Debug.Print path & FileName & vbTab & _
''              FileDateTime(path & FileName)   ' Include Modified Date
        ReDim Preserve aFileNames(FileCount)
        aFileNames(FileCount) = FileName
''#PFE#          Debug.Print "A:", FileCount, aFileNames(FileCount)
        FileCount = FileCount + 1
        FileName = Dir()  ' Get next file
    Wend

    ' If there are sub-directories...
    If nDir > 0 Then
        ' Recursively walk into them
        For i = 0 To nDir - 1
            aeFindFiles = aeFindFiles + aeFindFiles(path & dirNames(i) & "\", _
                                SearchStr, FileCount, DirCount, aFileNames())
        Next i
    End If

    If FileCount = 0 Then
        'MsgBox "No files found!"
        mblnFoundTheFile = False
    Else
        mblnFoundTheFile = True
    End If
    
PROC_EXIT:
    Exit Function

PROC_ERR:
    If Right(DirName, 4) = ".sys" Then
        Resume sysFileERRCont ' Known issue with pagefile.sys
    Else
        MsgBox "Erl=" & Erl & " Err=" & Err & " " & Err.Description, vbCritical, _
                    "aeFindFiles Error"
        Resume PROC_EXIT
    End If

End Function

Private Sub ListAllTheFiles(ByRef strTheNewLibFile As String)

    Debug.Print "ListAllTheFiles"

    Dim SearchPath As String
    Dim FindStr As String
    Dim FileSize As Long
    Dim NumFiles As Integer
    Dim NumDirs As Integer
    Dim aFNameResults() As String
    Dim i As Integer
    Dim strMessage As String

    On Error GoTo PROC_ERR

    mblnFoundTheFile = False
    DoCmd.Hourglass True
'      List2.Clear
    SearchPath = "C:\ae\testit\"
    FindStr = "adaeptdblib.mda.v*"
    FileSize = aeFindFiles(SearchPath, FindStr, NumFiles, NumDirs, aFNameResults())
    Debug.Print , NumFiles & " Files found in " & NumDirs + 1 & _
                                " Directories"
    Debug.Print , "Size of files found under " & SearchPath & " = " & _
            Format(FileSize, "#,###,###,##0") & " Bytes"
    DoCmd.Hourglass False

    If mblnFoundTheFile Then
        If UBound(aFNameResults()) > 1 Then
            For i = 0 To UBound(aFNameResults())
                Debug.Print "i = " & i, "aFNameResults(i) = " & aFNameResults(i)
                strMessage = strMessage & aFNameResults(i) & vbCrLf
            Next i
            MsgBox strMessage & vbCrLf & "Too many library files!" & vbCrLf & _
                "Please contact the system administrator.", vbCritical, "ListAllFiles"
        Else
            'MsgBox "The new library file is " & aFNameResults(i), vbInformation, "ListAllFiles"
            strTheNewLibFile = aFNameResults(i)
        End If
    End If

PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Erl=" & Erl & " Err=" & Err & " " & Err.Description, vbCritical, "ListAllTheFiles: Error"
    Resume PROC_EXIT

End Sub

Public Sub InstallNewLibrary()

    Debug.Print "InstallNewLibrary"

    Dim strLibFile As String
    
    On Error GoTo PROC_ERR

    ListAllTheFiles strLibFile
    Debug.Print , "The new lib file is: " & strLibFile
    'MsgBox "The new lib file is: " & strLibFile

    If mblnFoundTheFile Then

        ' Make backup copy of old library file
        If FileExists(gstrLocalLibPath & "adaeptdblib.mda.OLD") Then
            Kill gstrLocalLibPath & "adaeptdblib.mda.OLD"
        End If
        'MsgBox "gstrLocalLibPath & 'adaeptdblib.mda'=" & gstrLocalLibPath & "adaeptdblib.mda"
        'MsgBox "gstrLocalLibPath & 'adaeptdblib.mda.OLD'=" & gstrLocalLibPath & "adaeptdblib.mda.OLD"
        If FileExists(gstrLocalLibPath & "adaeptdblib.mda") Then
            Name gstrLocalLibPath & "adaeptdblib.mda" _
                As gstrLocalLibPath & "adaeptdblib.mda.OLD"
        End If
        If FileExists(gstrLocalLibPath & strLibFile) Then
            Name gstrLocalLibPath & strLibFile _
                  As gstrLocalLibPath & "adaeptdblib.mda"
        End If
    End If

PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Erl=" & Erl & " Err=" & Err & " " & Err.Description, vbCritical, "InstallNewLibrary"
    Resume PROC_EXIT

End Sub