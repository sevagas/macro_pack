VBA = \
r"""
' Exec process using WScript.Shell (asynchronous)
Sub WscriptExec(cmdLine As String)
    CreateObject("WScript.Shell").Run cmdLine, 0
End Sub
"""