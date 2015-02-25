Option Compare Database
Option Explicit

Public Function CurDir(Optional Drive As String)
CurDir = VBA.CurDir(Drive)
End Function

Public Function Environ(Expression)
Environ = VBA.Environ(Expression)
End Function