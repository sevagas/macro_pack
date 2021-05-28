VBA = \
r"""

Function Base64ToText(ByVal vCode)
    Dim oXML, oNode
    Dim tempString As String
    tempString = "Msxm"
    tempString = tempString & "l2.DO"
    tempString = tempString & "MDoc"
    tempString = tempString & "ument.3.0"
    Set oXML = CreateObject(tempString)
    Set oNode = oXML.CreateElement("base64")
    oNode.DataType = "bin.base64"
    oNode.Text = vCode
    Base64ToText = Stream_BinaryToString(oNode.nodeTypedValue)
    Set oNode = Nothing
    Set oXML = Nothing
End Function

'Stream_BinaryToString Function
'2003 Antonin Foller, http://www.motobit.com
'Binary - VT_UI1 | VT_ARRAY data To convert To a string
Private Function Stream_BinaryToString(Binary)
    Const adTypeText = 2
    Const adTypeBinary = 1
    
    'Create Stream object
    Dim BinaryStream 'As New Stream
    Dim tmpString As String
    tmpString = "ADO"
    tmpString = tmpString & "DB.St"
    tmpString = tmpString & "ream"
    Set BinaryStream = CreateObject(tmpString)
    
    'Specify stream type - we want To save binary data.
    BinaryStream.Type = adTypeBinary
    
    'Open the stream And write binary data To the object
    BinaryStream.Open
    BinaryStream.Write Binary
    
    'Change stream type To text/string
    BinaryStream.Position = 0
    BinaryStream.Type = adTypeText
    
    'Specify charset For the output text (unicode) data.
    BinaryStream.Charset = "us-ascii"
    
    'Open the stream And get text/string data from the object
    Stream_BinaryToString = BinaryStream.ReadText
    Set BinaryStream = Nothing
End Function

"""
