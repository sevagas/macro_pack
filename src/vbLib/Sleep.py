



VBA = \
r'''
#If VBA7 Then 
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else 
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If


Sub MySleep(nbSeconds)
    Sleep nbSeconds * 1000
End Sub

'''



VBA_XL = \
r'''
Sub MySleep(nbSeconds  As Integer)
    Dim dteWait
    dteWait = DateAdd("s", nbSeconds, Now())
    Do Until (Now() > dteWait)
        On Error Resume Next
        Application.Wait Now + TimeValue("0:00:01")
        On Error GoTo 0
    Loop
End Sub

'''

VBS = \
r'''
Sub MySleep(nbSeconds)
    WScript.Sleep nbSeconds * 1000
End Sub
'''

VBS_HTA = \
r'''
Sub MySleep(nbSeconds)
    Dim dteWait
    dteWait = DateAdd("s", nbSeconds, Now())
    Do Until (Now() > dteWait)
        On Error Resume Next
        CreateObject("WScript.Shell").Run "cmd /c ping localhost -n " & 1,0,True
        On Error GoTo 0
    Loop
End Sub

'''
