VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CExit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_component As New CComponent

Public Sub Initialise(ByRef rst As Recordset)
    Call m_component.Initialise(rst)
End Sub

Public Sub DrawLeft(ByRef roomPos As D3DVECTOR2)
    Call m_component.Draw(roomPos, False)
End Sub


Public Sub DrawRight(ByRef roomPos As D3DVECTOR2)
    Dim pos As D3DVECTOR2
    
    pos.X = roomPos.X + k_ROOM_WIDTH_PIX
    pos.Y = roomPos.Y
    Call m_component.Draw(pos, True)
End Sub

