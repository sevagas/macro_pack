VBA = \
"""
' Exec process using WMI
Private Sub WmiExec(targetPath As String)
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set objStartup = objWMIService.Get("Win32_ProcessStartup")
    Set objConfig = objStartup.SpawnInstance_
    Set objProcess = GetObject("winmgmts:\\.\root\cimv2:Win32_Process")
    errReturn = objProcess.Create(targetPath, Null, objConfig, intProcessID)
End Sub
"""