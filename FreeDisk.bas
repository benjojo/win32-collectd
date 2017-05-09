Attribute VB_Name = "FreeDisk"
Private Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type
Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, lpFreeBytesAvailableToCaller As LARGE_INTEGER, lpTotalNumberOfBytes As LARGE_INTEGER, lpTotalNumberOfFreeBytes As LARGE_INTEGER) As Long

Public Function GetDiskSpace(sDrive As String) As String
    Dim lResult As Long
    Dim liAvailable As LARGE_INTEGER
    Dim liTotal As LARGE_INTEGER
    Dim liFree As LARGE_INTEGER
    Dim dblAvailable As Double
    Dim dblTotal As Double
    Dim dblFree As Double
    If Right(sDrive, 1) <> "" Then sDrive = sDrive & ""
    'Determine the Available Space, Total Size and Free Space of a drive
    lResult = GetDiskFreeSpaceEx(sDrive, liAvailable, liTotal, liFree)
    
    'Convert the return values from LARGE_INTEGER to doubles
    dblAvailable = CLargeInt(liAvailable.lowpart, liAvailable.highpart)
    dblTotal = CLargeInt(liTotal.lowpart, liTotal.highpart)
    dblFree = CLargeInt(liFree.lowpart, liFree.highpart)
    
    'Display the results
    GetDiskSpace = "Available Space on " & sDrive & ":  " & dblAvailable & " bytes (" & _
                Format(dblAvailable / 1024 ^ 3, "0.00") & " G) " & vbCr & _
                "Total Space on " & sDrive & ":      " & dblTotal & " bytes (" & _
                Format(dblTotal / 1024 ^ 3, "0.00") & " G) " & vbCr & _
                "Free Space on " & sDrive & ":       " & dblFree & " bytes (" & _
                Format(dblFree / 1024 ^ 3, "0.00") & " G) "
End Function

Public Function GetTotalDiskSpace(sDrive As String) As String
    Dim lResult As Long
    Dim liAvailable As LARGE_INTEGER
    Dim liTotal As LARGE_INTEGER
    Dim liFree As LARGE_INTEGER
    Dim dblAvailable As Double
    Dim dblTotal As Double
    Dim dblFree As Double
    If Right(sDrive, 1) <> "" Then sDrive = sDrive & ""
    'Determine the Available Space, Total Size and Free Space of a drive
    lResult = GetDiskFreeSpaceEx(sDrive, liAvailable, liTotal, liFree)
    
    'Convert the return values from LARGE_INTEGER to doubles
    dblAvailable = CLargeInt(liAvailable.lowpart, liAvailable.highpart)
    dblTotal = CLargeInt(liTotal.lowpart, liTotal.highpart)
    dblFree = CLargeInt(liFree.lowpart, liFree.highpart)
    
    GetTotalDiskSpace = dblTotal
End Function


Public Function GetFreeDiskSpace(sDrive As String) As String
    Dim lResult As Long
    Dim liAvailable As LARGE_INTEGER
    Dim liTotal As LARGE_INTEGER
    Dim liFree As LARGE_INTEGER
    Dim dblAvailable As Double
    Dim dblTotal As Double
    Dim dblFree As Double
    If Right(sDrive, 1) <> "" Then sDrive = sDrive & ""
    'Determine the Available Space, Total Size and Free Space of a drive
    lResult = GetDiskFreeSpaceEx(sDrive, liAvailable, liTotal, liFree)
    
    'Convert the return values from LARGE_INTEGER to doubles
    dblAvailable = CLargeInt(liAvailable.lowpart, liAvailable.highpart)
    dblTotal = CLargeInt(liTotal.lowpart, liTotal.highpart)
    dblFree = CLargeInt(liFree.lowpart, liFree.highpart)
    
    GetFreeDiskSpace = dblFree
End Function

Private Function CLargeInt(Lo As Long, Hi As Long) As Double
    'This function converts the LARGE_INTEGER data type to a double
    Dim dblLo As Double, dblHi As Double
    
    If Lo < 0 Then
        dblLo = 2 ^ 32 + Lo
    Else
        dblLo = Lo
    End If
    
    If Hi < 0 Then
        dblHi = 2 ^ 32 + Hi
    Else
        dblHi = Hi
    End If
    CLargeInt = dblLo + dblHi * 2 ^ 32
End Function

