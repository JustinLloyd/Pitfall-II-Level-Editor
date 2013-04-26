VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBackground"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_component As New CComponent

Public Sub Initialise(ByRef rst As Recordset)
    Call m_component.Initialise(rst)
End Sub

Public Sub Draw(ByRef roomPos As D3DVECTOR2)
    Call m_component.Draw(roomPos)
End Sub
