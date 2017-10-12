Private Sub toto()
    MsgBox "Hello Word from macropack"
End Sub


Private Sub testMacroXl()
    Application.Run "ThisWorkbook.toto"
End Sub

Private Sub testMacroWd()
    toto
End Sub

' triggered when Word/Powerpoint generator is used 
Sub AutoOpen()
    testMacroWd
End Sub


' triggered when Ecel generator is used
Sub Workbook_Open()
    testMacroXl
End Sub
