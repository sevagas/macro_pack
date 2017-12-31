VBA = \
r"""
' Exec process using WMI
Sub WscriptExec(targetPath As String)
    CreateObject("WScript.Shell").Run targetPath, 0
End Sub
"""