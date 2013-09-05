Attribute VB_Name = "FileUtils"
' create folder if parent dose not exist.
' equals mkdir -p under *NIX
Function MkDirParents(path As String)
  On Error GoTo Error_Handle
  
  Dim SubPath As String
  Dim PathParts() As String
  Dim Part As String
  
  Dim Temp As String
  
  'ensure there is no double backslashes.
  PathParts = Split(Replace(path, "\\", "\"), "\")
  
  For i = LBound(PathParts) To UBound(PathParts)
    Part = PathParts(i)
    If Right(Part, 1) = ":" Then
      SubPath = Part
    Else
      SubPath = SubPath & "\" & Part
      Temp = dir(SubPath, vbDirectory)
      If Temp = "" Then
        ' folder not exist,create it.
        MkDir SubPath
      End If

    End If
  Next i
  
  MkDirParents = True
  Exit Function
Error_Handle:
  MkDirParents = False

End Function

