VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CVCSImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private m_romData() As Byte

Private Sub Class_Initialize()
    m_romData = LoadResData(101, "CUSTOM")
    DumpROM
End Sub

Public Sub SetByte(ByVal addr As Long, ByVal val As Byte)
    Debug.Assert addr >= LBound(m_romData) And addr <= UBound(m_romData)
    m_romData(addr) = val
End Sub

Public Function GetData() As Byte()
    GetData = m_romData
End Function

Private Function DumpROM()
    Dim addr As Long
    Dim x As Integer
    Dim y As Integer
    Dim lowNibble As Byte
    
    Debug.Print "ROM Dump"
    addr = 3271
    For y = 0 To 31
        Debug.Print "$"; HexInt(addr); " ";
        For x = 0 To 7
            lowNibble = m_romData(addr) And 15
            If lowNibble = 6 Or lowNibble = 7 Or lowNibble = 15 Then
                Debug.Assert False
            End If
            
            Debug.Print HexByte(m_romData(addr)); " ";
            addr = addr + 1
            
        Next
        Debug.Print
    Next
End Function
