Private Sub toto()
    MsgBox "Hello Word from macropack"
End Sub


Private Sub testMacro()
    Application.Run "ThisWorkbook.toto"
End Sub



' triggered when document is opened
Sub Workbook_Open()
    testMacro
End Sub
