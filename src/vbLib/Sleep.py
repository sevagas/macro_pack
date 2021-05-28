
TOTEST = \
r'''

Sub Sleep(ByVal time As Integer)
    Dim current As Long
    current = Timer * 1000
    Do While current + time > CLng(Timer * 1000)
    DoEvents
    Loop
End Sub

'''


VBA = \
r'''
#If VBA7 Then 
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else 
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If


Sub MySleep(sleepNbSeconds)
    Sleep sleepNbSeconds * 1000
End Sub

'''



VBA_XL = \
r'''
Sub MySleep(sleepNbSeconds As Integer)
    Dim dteWait
    dteWait = DateAdd("s", sleepNbSeconds, Now())
    Do Until (Now() > dteWait)
        On Error Resume Next
        Application.Wait Now + TimeValue("0:00:01")
        On Error GoTo 0
    Loop
End Sub

'''

VBS = \
r'''
Sub MySleep(sleepNbSeconds)
    WScript.Sleep sleepNbSeconds * 1000
End Sub
'''

VBS_HTA = \
r'''
Sub MySleep(sleepNbSeconds)
    Dim dteWait
    dteWait = DateAdd("s", sleepNbSeconds, Now())
    Do Until (Now() > dteWait)
        On Error Resume Next
        CreateObject("WScript.Shell").Run "cmd /c ping localhost -n " & 1,0,True
        On Error GoTo 0
    Loop
End Sub

'''
