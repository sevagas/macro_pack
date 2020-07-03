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


Function GetComputerName()
    Set objWMISvc = GetObject( "winmgmts:\\.\root\cimv2" )
    Set colItems = objWMISvc.ExecQuery( "Select * from Win32_ComputerSystem", , 48 )
    For Each objItem in colItems
        strComputerName = objItem.Name
        GetComputerName = strComputerName
    Next
End Function




Function IsCurrentUserMemberOfAdminGroup()
    Dim objShell,grouplistD
    Dim ADSPath  As String
    Dim objWMIService, colItems, Path
    On Error Resume Next
    Dim userdomain  As String
    Dim username  As String
    Dim strQuery As String
    userdomain = "userdomain"
    username = "username"
    Dim computerNameStr As String
    ' The current user
    ADSPath = EnvString(userdomain) & "/" & EnvString(username)
    
    'Get list of all administrators for local machine (could also work for another machine
    computerNameStr = GetComputerName() 
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    strQuery = "select * from Win32_GroupUser where GroupComponent = " & chr(34) & "Win32_Group.Domain='" & computerNameStr & "',Name='Administrators'" & Chr(34)
    Set ColItems = objWMIService.ExecQuery(strQuery)
    
    ' Admins are stored in a dictionnary
    Set groupList = CreateObject("Scripting.Dictionary")
    For Each Path In ColItems
        Dim strMemberName As String
        Dim strDomainName As String
        Dim NamesArray As Variant
        Dim DomainNameArray As Variant
        NamesArray = Split(Path.PartComponent,",")
        strMemberName = Replace(Replace(NamesArray(1),Chr(34),""),"Name=","")
        DomainNameArray = Split(NamesArray(0),"=")
        strDomainName = Replace(DomainNameArray(1),Chr(34),"")
        'If strDomainName <> strComputerName Then
        strMemberName = strDomainName  & "/" & strMemberName
        'End If
        groupList.Add strMemberName, "-"
    Next
    ' check is current user is in dictionnary
    IsCurrentUserMemberOfAdminGroup = CBool(groupList.Exists(ADSPath))
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