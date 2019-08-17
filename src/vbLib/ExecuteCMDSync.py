

VBA = \
r"""

Function ExecuteCmdSync(targetPath As String)
    'Run a shell command, returning the output as a string'
    ' Using a hidden window, pipe the output of the command to the CLIP.EXE utility...
    ' Necessary because normal usage with oShell.Exec("cmd.exe /C " & sCmd) always pops a windows
    Dim instruction As String
    instruction = "cmd.exe /c " & targetPath & " | clip"
    
    On Error Resume Next
    Err.Clear
    CreateObject("WScript.Shell").Run instruction, 0, True
    On Error Goto 0
    
    ' Read the clipboard text using htmlfile object
    ExecuteCmdSync = CreateObject("htmlfile").ParentWindow.ClipboardData.GetData("text")

End Function


"""

