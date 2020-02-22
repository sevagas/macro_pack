VBA = \
r"""
 'Create A  Text and fill it
 ' Will overwrite existing file
Sub CreateBinFile(FilePath As String, bytes)
    Dim binaryStream
    Set binaryStream = CreateObject("ADODB.Stream")
    binaryStream.Type = 1 ' 1 = TypeBinary
    'Open the stream and write binary data
    binaryStream.Open
    binaryStream.Write bytes
    'Save binary data to disk
    binaryStream.SaveToFile FilePath, 2 ' 2 = ForWriting
End Sub

"""