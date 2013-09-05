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

' list files into a array
' path : target flder
' filter : dir filter ,for example "\*.doc"
' why this function? because you can not use dir nested.
Function ListFiles(path As String, filter As String)
  
  Dim buf As String
  Dim size  As Integer
  Dim files() As String
  
  ' set array's size to 0
  size = 0
  
  ' get first file
  buf = dir(path & filter)
  
  Do While buf <> ""
    
    'redim array
    ReDim Preserve files(size)
    
    'push file to array
    files(size) = buf
    
    size = size + 1
    
    'get next file
    buf = dir()
  Loop
  
  ListFiles = files

End Function

