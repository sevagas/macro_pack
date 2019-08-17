

VBA = \
r"""

Sub ExecuteCmdAsync(targetPath As String)
    On Error Resume Next
    Err.Clear
    wimResult = WmiExec(targetPath)
    If Err.Number <> 0 Or wimResult <> 0 Then
        Err.Clear
        WscriptExec targetPath
    End If
    On Error Goto 0
End Sub
"""