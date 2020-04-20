"""
Get the windows OS version
"""


r"""

#If VBA7 Then 
    Declare PtrSafe Function RtlGetVersion Lib "NTDLL" (ByRef lpVersionInformation As Long) As Long
#Else 
    Declare Function RtlGetVersion Lib "NTDLL" (ByRef lpVersionInformation As Long) As Long
#End If

Public Function GetOSVersion() As String
    Dim tOSVw(&H54) As Long
    tOSVw(0) = &H54 * &H4
    Call RtlGetVersion(tOSVw(0))
    'GetOSVersion = Join(Array(tOSVw(1), tOSVw(2), tOSVw(3)), ".")
    GetOSVersion = VersionToName(Join(Array(tOSVw(1), tOSVw(2)), "."))
End Function

Private Function VersionToName(ByRef sVersion As String) As String
    Select Case sVersion
        Case "5.1": VersionToName = "Windows XP"
        Case "5.3": VersionToName = "Windows 2003 (SERVER)"
        Case "6.0": VersionToName = "Windows Vista"
        Case "6.1": VersionToName = "Windows 7"
        Case "6.2": VersionToName = "Windows 8"
        Case "6.3": VersionToName = "Windows 8.1"
        Case "10.0": VersionToName = "Windows 10"
        Case Else: VersionToName = "Unknown"
    End Select
End Function


"""

VBA = \
r'''

Public Function GetOSVersion() As String
    Dim prodType As String
    Dim version As String
    Dim desktopProductType As String
    desktopProductType = "1"
    For Each objItem in GetObject("winmgmts://./root/cimv2").ExecQuery("Select * from Win32_OperatingSystem",,48)
        version = objItem.Version
        prodType = objItem.ProductType & ""
    Next
 
    Select Case Left(version, Instr(version, ".") + 1)
    Case "10.0"
        If (prodType = desktopProductType) Then
            GetOSVersion = "Windows 10"
        Else
            GetOSVersion = "Windows Server 2016"
        End If
    Case "6.3"
        If (prodType = desktopProductType) Then
            GetOSVersion = "Windows 8.1"
        Else
            GetOSVersion = "Windows Server 2012 R2"
        End If
    Case "6.2"
        If (prodType = desktopProductType) Then
            GetOSVersion = "Windows 8"
        Else
            GetOSVersion = "Windows Server 2012"
        End If
    Case "6.1"
        If (prodType = desktopProductType) Then
            GetOSVersion = "Windows 7"
        Else
            GetOSVersion = "Windows Server 2008 R2"
        End If
    Case "6.0"
        If (prodType = desktopProductType) Then
            GetOSVersion = "Windows Vista"
        Else
            GetOSVersion = "Windows Server 2008"
        End If
    Case "5.2"
        If (prodType = desktopProductType) Then
            GetOSVersion = "Windows XP 64-Bit Edition"
        ElseIf (Left(Version, 5) = "5.2.3") Then
            GetOSVersion = "Windows Server 2003 R2"
        Else
            GetOSVersion = "Windows Server 2003"
        End If
    Case "5.1"
        GetOSVersion = "Windows XP"
    Case "5.0"
        GetOSVersion = "Windows 2000"
    End Select
End Function

'''