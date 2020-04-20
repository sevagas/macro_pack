VBA = \
"""
 'Create A  Text and fill it
 ' Will overwrite existing file
Sub CreateTxtFile(FilePath As String, FileContent As String)
   
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim Fileout As Object
    Set Fileout = fso.CreateTextFile(FilePath, True)
    Fileout.Write FileContent
    Fileout.Close

End Sub
"""
