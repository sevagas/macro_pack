VBA = \
r"""

Sub WriteBytes(objFile, strBytes)
    Dim aNumbers
    Dim iIter

    aNumbers = split(strBytes)
    for iIter = lbound(aNumbers) to ubound(aNumbers)
        objFile.Write Chr(aNumbers(iIter))
    next
End Sub

"""