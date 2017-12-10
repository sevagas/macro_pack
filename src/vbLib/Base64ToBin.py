VBA = \
r"""

Private Function Base64ToBin(base64)
    Dim DM, EL
    Set DM = CreateObject("Microsoft.XMLDOM")
    ' Create temporary node with Base64 data type
    Set EL = DM.createElement("tmp")
    EL.DataType = "bin.base64"
    ' Set encoded String, get bytes
    EL.Text = base64
    Base64ToBin = EL.NodeTypedValue
End Function

"""