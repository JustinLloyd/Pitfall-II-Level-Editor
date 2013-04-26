Attribute VB_Name = "MathFunctions"
Option Explicit

Public Function HexBytesFromInteger(ByVal val As Long) As String
    Dim b As Long
    Dim s As String
    
    s = ""
    b = val And &HFF
    s = s & "$" & HexByte(b) & ","
    b = val And &HFF00
    b = RShift(b, 8)
    s = s & "$" & HexByte(b)
'    b = val And &HFF0000
'    b = RShift(b, 16)
'    s = s & "$" & HexByte(b) & ","
'    b = val And &HFF000000
'    s = s & "$" & HexByte(b)
'    b = RShift(b, 24)
    
    HexBytesFromInteger = s
End Function

Public Function HexByte(ByVal val As Long) As String
    Dim s As String
    
    s = Hex$(val)
    While Len(s) < 2
        s = "0" + s
    Wend
    
    HexByte = s
End Function

Public Function HexInt(ByVal val As Long) As String
    Dim s As String
    
    s = Hex$(val)
    While Len(s) < 4
        s = "0" + s
    Wend
    
    HexInt = s

End Function

Public Function IntMax(ByVal a As Integer, ByVal b As Integer) As Integer
    If a > b Then
        IntMax = a
    Else
        IntMax = b
    End If
End Function

Public Function IntMin(ByVal a As Integer, ByVal b As Integer) As Integer
    If a < b Then
        IntMin = a
    Else
        IntMin = b
    End If
End Function

Public Function LShift(ByVal operand As Long, ByVal bits As Integer) As Long
    If bits = 0 Then
        LShift = operand
    Else
        LShift = operand * (2 ^ bits)
    End If
End Function

Public Function RShift(ByVal operand As Long, ByVal bits As Integer) As Long
    If bits = 0 Then
        RShift = operand
    Else
        RShift = operand \ (2 ^ bits)
    End If
End Function


