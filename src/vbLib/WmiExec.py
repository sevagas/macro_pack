VBA = \
r"""
' Exec process using WMI
Function WmiExec(targetPath As String) As Integer
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set objStartup = objWMIService.Get("Win32_ProcessStartup")
    Set objConfig = objStartup.SpawnInstance_
    Set objProcess = GetObject("winmgmts:\\.\root\cimv2:Win32_Process")
    WmiExec = objProcess.Create(targetPath, Null, objConfig, intProcessID)
End Function
"""