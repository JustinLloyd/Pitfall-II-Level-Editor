VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHazard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_component As New CComponent
Private m_movementType As Integer

Public Sub Initialise(ByRef rst As Recordset)
    m_movementType = rst!MovementType
    Call m_component.Initialise(rst)
End Sub

Public Sub Draw(ByRef roomPos As D3DVECTOR2)
    Dim dx As Long
    Dim flightPathIdx As Integer
    Dim rotCentre As D3DVECTOR2
    Dim pos As D3DVECTOR2
    Dim frameIdx As Integer
    Dim flightPathLen As Integer
    Dim hFlip As Boolean
    
    hFlip = False
    frameIdx = g_frameCount Mod m_component.FrameCount
    flightPathLen = UBound(g_flightPath) + 1
    pos.x = 0
    pos.y = 0
    Select Case m_movementType
        Case 1 ' condor
            If ((g_frameCount \ (k_ROOM_WIDTH_PIX + 16)) And 1) = 0 Then
                pos.x = (k_ROOM_WIDTH_PIX + 8) - (g_frameCount Mod (k_ROOM_WIDTH_PIX + 16))
            Else
                pos.x = -8 + g_frameCount Mod (k_ROOM_WIDTH_PIX + 16)
                hFlip = True
            End If
            
            flightPathIdx = (g_frameCount \ 2&) Mod flightPathLen
            pos.y = g_flightPath(flightPathIdx).y
        
        Case 2 ' frog
        Case 3 ' scorpion
            pos.x = -8 + (g_frameCount Mod (k_ROOM_WIDTH_PIX + 16))
        
        Case 4 ' ant
        Case 5 ' eel
            frameIdx = g_frameCount Mod (m_component.FrameCount / 2)
            If Int(Rnd * 100) > 50 Then frameIdx = frameIdx + 2
            pos.x = g_frameCount * 2& Mod k_ROOM_WIDTH_PIX
            'pos.y = k_CREATURE_EEL_Y
        
        Case 6 ' bat
            pos.x = g_frameCount Mod k_ROOM_WIDTH_PIX
            flightPathIdx = g_frameCount Mod flightPathLen
            pos.y = g_flightPath(flightPathIdx).y
        
'    ElseIf room.Creature = k_CREATURE_EEL Then
'        animFrame = frameCount Mod 2&
'        If Int(Rnd * 100) > 50 Then animFrame = animFrame + 2
'        dx = frameCount * 2& Mod 160&
'        dy = k_CREATURE_EEL_Y
'    ElseIf room.Creature = k_CREATURE_FIRE_ANT Then
'        animFrame = frameCount Mod 2&
'        dx = frameCount * 2& Mod 320&
'        dx = Abs(dx - 160)
'        dy = k_CREATURE_FIRE_ANT_Y
'    ElseIf room.Creature = k_CREATURE_FROG Then
'        animFrame = frameCount Mod 2&
'        dx = 80
'        dy = 30
'    Else
'        animFrame = 0
'        dx = 0
'        dy = 0
'    End If
        Case Else
            pos.x = 0
            pos.y = 0

    End Select
    
    pos.x = pos.x + roomPos.x
    pos.y = pos.y + roomPos.y
    
    rotCentre.x = 0
    rotCentre.y = 0

    Call m_component.DrawFrame(frameIdx, pos, hFlip)
    
End Sub

