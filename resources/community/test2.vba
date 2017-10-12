Private Sub toto()
    MsgBox "Hello from <<<TEMPLATE>>>" & vbCrLf & "Remember to always be careful when you enable MS Office macros." & vbCrLf & "Have a nice day!"
End Sub

Private Sub testMacroXl()
    Application.Run "ThisWorkbook.toto"
End Sub

Private Sub testMacroWd()
    toto
End Sub

' triggered when Word/PowerPoint generator is used 
Sub AutoOpen()
    testMacroWd
End Sub

' triggered when Excel generator is used
Sub Workbook_Open()
    testMacroXl
End Sub
