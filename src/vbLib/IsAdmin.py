VBA = \
r'''


Function IsAdmin()
    On Error Resume Next
    CreateObject("WScript.Shell").RegRead("HKEY_USERS\S-1-5-19\Environment\TEMP")
    if Err.number = 0 Then 
        IsAdmin = True
    else
        IsAdmin = False
    end if
    Err.Clear
    On Error goto 0
End Function



'This function checks to see if the passed group name contains the current
' user as a member. Returns True or False
Function IsMember(groupName)
    Dim objShell,grouplistD,ADSPath,userPath,listGroup
    On Error Resume Next
    
    set objShell = CreateObject( "WScript.Shell" )
    If IsEmpty(groupListD) then
        Set groupListD = CreateObject("Scripting.Dictionary")
        groupListD.CompareMode = 1
        ADSPath = EnvString("userdomain") & "/" & EnvString("username")
        Set userPath = GetObject("WinNT://" & ADSPath & ",user")
        For Each listGroup in userPath.Groups
            groupListD.Add listGroup.Name, "-"
        Next
    End if
    IsMember = CBool(groupListD.Exists(groupName))
    ' Clean up
    Set objShell = Nothing
End Function



'This function returns a particular environment variable's value.
' for example, if you use EnvString("username"), it would return
' the value of %username%.
Function EnvString(variable)
    Dim objShell    
    set objShell = CreateObject( "WScript.Shell" )
    variable = "%" & variable & "%"
    EnvString = objShell.ExpandEnvironmentStrings(variable)
    ' Clean up
    Set objShell = Nothing
End Function

 



'''