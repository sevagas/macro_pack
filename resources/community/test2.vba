Private Sub toto()
    MsgBox "Hello Word from macropack"
End Sub


Private Sub testMacro()
    Application.Run "ThisWorkbook.toto"
End Sub


' triggered when Word/Powerpoint generator is used 
Sub AutoOpen()
    testMacro
End Sub


' triggered when Ecel generator is used
Sub Workbook_Open()
    testMacro
End Sub
