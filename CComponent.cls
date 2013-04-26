VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CComponent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_name As String
Private m_frameCount As Integer
Private m_graphic() As Direct3DTexture8
Private m_scaling() As D3DVECTOR2
Private m_pos As D3DVECTOR2           ' the location of this actor on screen
Private m_sprite As D3DXSprite

Public Sub Initialise(ByRef rst As Recordset)
    Dim i As Integer
    Dim texturePath As String

    m_name = rst!ListBoxName
    m_frameCount = rst!FrameCount
    m_pos.X = rst!X
    m_pos.Y = rst!Y

    texturePath = App.path & "\NewBitmaps\"
    ReDim m_graphic(0 To m_frameCount - 1)
    ReDim m_scaling(0 To m_frameCount - 1)
    For i = 0 To m_frameCount - 1
        Set m_graphic(i) = g_d3dObject.LoadTexture(texturePath & rst.Fields("Frame" & Trim(i)) & ".bmp", &HFF00FF, m_scaling(i))
    Next

    Set m_sprite = g_d3dX.CreateSprite(g_d3dDevice)
End Sub


Public Sub Draw(ByRef pos As D3DVECTOR2, Optional ByVal hFlip As Boolean)
    Dim rotCentre As D3DVECTOR2
    Dim newPos As D3DVECTOR2
    Dim frameIdx As Integer
    
    frameIdx = g_frameCount Mod m_frameCount
    Call DrawFrame(frameIdx, pos, hFlip)
'    m_sprite.Draw m_graphic(frameIdx), ByVal 0&, m_scaling(frameIdx), rotCentre, ByVal 0#, newPos, &HFFFFFFFF
End Sub

Public Sub DrawFrame(ByVal frameIdx As Integer, ByRef pos As D3DVECTOR2, Optional ByVal hFlip As Boolean)
    Dim rotCentre As D3DVECTOR2
    Dim newPos As D3DVECTOR2
    Dim scaling As D3DVECTOR2
    
    rotCentre.X = 0
    rotCentre.Y = 0

    newPos.X = pos.X + m_pos.X
    newPos.Y = pos.Y + m_pos.Y
    
    scaling.Y = m_scaling(frameIdx).Y
    If hFlip Then
        scaling.X = -m_scaling(frameIdx).X
    Else
        scaling.X = m_scaling(frameIdx).X
    End If
    
    m_sprite.Draw m_graphic(frameIdx), ByVal 0&, scaling, rotCentre, ByVal 0#, newPos, &HFFFFFFFF
End Sub


Public Property Get XPos() As Single
    XPos = m_pos.X
End Property

Public Property Get YPos() As Single
    YPos = m_pos.Y
End Property

Public Property Get FrameCount() As Integer
    FrameCount = m_frameCount
End Property

'Public Property Get Texture(ByVal idx As Integer) As Direct3DTexture8
'    Set Texture = m_graphic(idx)
'End Property

