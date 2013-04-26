VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_hasChanged As Boolean
Private m_groundType As Integer
Private m_floorType As Integer
Private m_exitLeftType  As Integer
Private m_exitRightType As Integer
Private m_backgroundType As Integer
Private m_itemType As Integer
Private m_hazardType As Integer
Private m_featureType As Integer
Private m_linkToLevel As Integer                          ' level # that this room links to, 0 if no link
Private m_linkToRoom As Integer                           ' room # that this room links to, must be set if m_linkToLevel is set
Private m_row As Integer
Private m_col As Integer

Public Property Let LinkToRoom(ByVal roomNum As Integer)
    If roomNum <> m_linkToRoom Then
        m_linkToRoom = roomNum
        FlagRoomChange
    End If
End Property

Public Property Get LinkToRoom() As Integer
    LinkToRoom = m_linkToRoom
End Property

Public Property Let LinkToLevel(ByVal levelNum As Integer)
    If levelNum <> m_linkToLevel Then
        m_linkToLevel = levelNum
        FlagRoomChange
    End If
End Property

Public Property Get LinkToLevel() As Integer
    LinkToLevel = m_linkToLevel
End Property

Public Property Get HasChanged() As Boolean
    HasChanged = m_hasChanged
End Property

Public Sub ClearChange()
    m_hasChanged = False
End Sub

Private Sub FlagRoomChange()
    m_hasChanged = True
End Sub

Public Property Let Feature(ByVal featureType As Integer)
    Debug.Assert featureType >= 0 And featureType <= UBound(g_featureType)
    If featureType <> m_featureType Then
        m_featureType = featureType
        FlagRoomChange
    End If

End Property

Public Property Get Feature() As Integer
    Feature = m_featureType
End Property

Public Property Let Item(ByVal itemType As Integer)
    Debug.Assert itemType >= 0 And itemType <= UBound(g_itemType)
    If itemType <> m_itemType Then
        m_itemType = itemType
        FlagRoomChange
    End If

End Property

Public Property Get Item() As Integer
    Item = m_itemType
End Property

Public Property Let Hazard(ByVal hazardType As Integer)
    Debug.Assert hazardType >= 0 And hazardType <= UBound(g_hazardType)
    If hazardType <> m_hazardType Then
        m_hazardType = hazardType
        FlagRoomChange
    End If

End Property

Public Property Get Hazard() As Integer
    Hazard = m_hazardType
End Property

Public Property Let Ground(ByVal groundType As Integer)
    Debug.Assert groundType >= 0 And groundType <= UBound(g_groundType)
    If groundType <> m_groundType Then
        m_groundType = groundType
        FlagRoomChange
    End If
    
End Property

Public Property Get Ground() As Integer
    Ground = m_groundType
End Property

Public Property Let Floor(ByVal floorType As Integer)
    Debug.Assert floorType >= 0 And floorType <= UBound(g_floorType)
    If floorType <> m_floorType Then
        m_floorType = floorType
        FlagRoomChange
    End If

End Property

Public Property Get Floor() As Integer
    Floor = m_floorType
End Property

Public Property Let ExitLeft(ByVal exitType As Integer)
    Debug.Assert exitType >= 0 And exitType <= UBound(g_exitType)
    If exitType <> m_exitLeftType Then
        m_exitLeftType = exitType
        FlagRoomChange
    End If
    
End Property

Public Property Get ExitLeft() As Integer
    ExitLeft = m_exitLeftType
End Property

Public Property Let ExitRight(ByVal exitType As Integer)
    Debug.Assert exitType >= 0 And exitType <= UBound(g_exitType)
    If exitType <> m_exitRightType Then
        m_exitRightType = exitType
        FlagRoomChange
    End If
    
End Property

Public Property Get ExitRight() As Integer
    ExitRight = m_exitRightType
End Property

Public Property Let Background(ByVal backgroundType As Integer)
    Debug.Assert backgroundType >= 0 And backgroundType <= UBound(g_backgroundType)
    If backgroundType <> m_backgroundType Then
        m_backgroundType = backgroundType
        FlagRoomChange
    End If
    
End Property

Public Property Get Background() As Integer
    Background = m_backgroundType
End Property

Public Property Let column(ByVal newCol As Integer)
    m_col = newCol
End Property

Public Property Get column() As Integer
    column = m_col
End Property

Public Property Let row(ByVal newRow As Integer)
    m_row = newRow
End Property

Public Property Get row() As Integer
    row = m_row
End Property

Private Sub Class_Initialize()
    InitialiseRoom
End Sub

Public Sub SerialIn(ByVal id As String, ByRef proj As CProject)
    Dim prop As PropertyBag
    Dim result As Boolean

    Set prop = New PropertyBag
    result = proj.GetContent(id, prop)
    Debug.Assert result <> False
    m_exitLeftType = prop.ReadProperty("ExitLeft", 0)
    m_exitRightType = prop.ReadProperty("ExitRight", 0)
    m_backgroundType = prop.ReadProperty("Background", 0)
    m_floorType = prop.ReadProperty("Floor", 0)
    m_groundType = prop.ReadProperty("Ground", 0)
    m_itemType = prop.ReadProperty("Item", 0)
    m_hazardType = prop.ReadProperty("Hazard", 0)
    m_featureType = prop.ReadProperty("Feature", 0)
    m_linkToLevel = prop.ReadProperty("LinkToLevel", 0)
    m_linkToRoom = prop.ReadProperty("LinkToRoom", 0)
End Sub

Public Sub SerialOut(ByVal id As String, ByRef proj As CProject)
    Dim prop As PropertyBag
    Dim result As Boolean

    Set prop = New PropertyBag
    prop.WriteProperty "ExitLeft", m_exitLeftType
    prop.WriteProperty "ExitRight", m_exitRightType
    prop.WriteProperty "Background", m_backgroundType
    prop.WriteProperty "Floor", m_floorType
    prop.WriteProperty "Ground", m_groundType
    prop.WriteProperty "Item", m_itemType
    prop.WriteProperty "Hazard", m_hazardType
    prop.WriteProperty "Feature", m_featureType
    prop.WriteProperty "LinkToLevel", m_linkToLevel
    prop.WriteProperty "LinkToRoom", m_linkToRoom
    result = proj.AddContent(id, prop)
    Debug.Assert result <> False
End Sub

Public Function ExportSource() As String
'    Dim exportString As String
'    Dim featureFlags As Long
'
'    exportString = ""
'    exportString = exportString & "$" & HexByte(m_floorType) & ","
'    exportString = exportString & "$" & HexByte(m_backgroundType) & ","
'    exportString = exportString & "$" & HexByte(m_exitLeftType) & ","
'    exportString = exportString & "$" & HexByte(m_exitRightType) & ","
'    exportString = exportString & "$" & HexByte(m_HazardType) & ","
'    exportString = exportString & "$" & HexByte(m_itemType) & ","
'
'    exportString = exportString & HexBytesFromInteger(featureFlags)
'    ExportSource = exportString
End Function

Public Sub Copy(ByRef srcRoom As CRoom)
    FlagRoomChange
    m_floorType = srcRoom.Floor
    m_groundType = srcRoom.Ground
    m_exitLeftType = srcRoom.ExitLeft
    m_exitRightType = srcRoom.ExitRight
    m_backgroundType = srcRoom.Background
    m_itemType = srcRoom.Item
    m_hazardType = srcRoom.Hazard
    m_featureType = srcRoom.Feature
    m_linkToLevel = srcRoom.LinkToLevel
    m_linkToRoom = srcRoom.LinkToRoom
    m_row = srcRoom.row
    m_col = srcRoom.column
End Sub

Public Sub Reset()
    InitialiseRoom
    FlagRoomChange
End Sub

Private Sub InitialiseRoom()
    m_hasChanged = False
    m_floorType = 0
    m_groundType = 0
    m_backgroundType = 0
    m_exitLeftType = 0
    m_exitRightType = 0
    m_itemType = 0
    m_hazardType = 0
    m_featureType = False
    m_linkToLevel = 0
    m_linkToRoom = 0
End Sub

Public Sub Draw(ByRef pos As D3DVECTOR2)
    If m_backgroundType >= LBound(g_backgroundType) And m_backgroundType <= UBound(g_backgroundType) Then
        Call g_backgroundType(m_backgroundType).Draw(pos)
    End If
    
    If m_groundType > 0 And m_groundType <= UBound(g_groundType) Then
        Call g_groundType(m_groundType).Draw(pos)
    End If
    
    If m_floorType > 0 And m_floorType <= UBound(g_floorType) Then
        Call g_floorType(m_floorType).Draw(pos)
    End If
    
    If m_exitLeftType > 0 And m_exitLeftType <= UBound(g_exitType) Then
        Call g_exitType(m_exitLeftType).DrawLeft(pos)
    End If
    
    If m_exitRightType > 0 And m_exitRightType <= UBound(g_exitType) Then
        Call g_exitType(m_exitRightType).DrawRight(pos)
    End If
    
    If m_featureType >= LBound(g_featureType) And m_featureType <= UBound(g_featureType) Then
        Call g_featureType(m_featureType).Draw(pos)
    End If
    
    If m_hazardType > 0 Then
        Call g_hazardType(m_hazardType).Draw(pos)
    End If
    
    If m_itemType > 0 Then
        Call g_itemType(m_itemType).Draw(pos)
    End If
    
End Sub
