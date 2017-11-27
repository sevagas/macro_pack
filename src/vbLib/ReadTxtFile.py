VBA = \
"""
' Read the content of a text file and return it
' Return empty if file does not exist
Function ReadTxtFile(FilePath As String) As String

    Dim fileNo As Integer
    Dim FileContent As String
    
    ' Exit if file does not exist
    If Len(Dir$(FilePath)) = 0 Then
        Exit Function
    End If
    
    'Determine the next file number available for use by the FileOpen function
      fileNo = FreeFile()
    
    'Open the text file
      Open FilePath For Input As fileNo
    
    'Store file content inside a variable
      Do While Not EOF(fileNo)
        Line Input #fileNo, FileContent
      Loop
      
    'Close Text File
      Close #fileNo
      ReadTxtFile = FileContent
End Function
"""
