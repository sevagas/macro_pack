VBA = \
r"""
' Exec process using WScript.Shell (asynchronous)
Sub WscriptExec(targetPath As String)
    CreateObject("WScript.Shell").Run targetPath, 0
End Sub
"""