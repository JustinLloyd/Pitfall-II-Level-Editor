VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_undoCollection As Collection
Private m_mapWidth As Integer
Private m_mapHeight As Integer
Private m_startRoom As Integer              ' first room that player starts in on this map
Private m_hasChanged As Boolean

Private m_room(0 To k_MAX_MAP_WIDTH - 1, 0 To k_MAX_MAP_HEIGHT - 1) As CRoom

Public Sub ClearChange()
    Dim col As Integer
    Dim row As Integer
    
    m_hasChanged = False
    For col = 0 To k_MAX_MAP_WIDTH - 1
        For row = 0 To k_MAX_MAP_HEIGHT - 1
            m_room(col, row).ClearChange
        Next
    Next
End Sub

Public Function HasChanged() As Boolean
    Dim col As Integer
    Dim row As Integer
    
    If m_hasChanged Then
        HasChanged = True
        Exit Function
    End If
    
    For col = 0 To k_MAX_MAP_WIDTH - 1
        For row = 0 To k_MAX_MAP_HEIGHT - 1
            If m_room(col, row).HasChanged Then
                HasChanged = True
                Exit Function
            End If
        Next
    Next
    
    HasChanged = False
End Function

Private Sub FlagMapChange()
    m_hasChanged = True
End Sub

Public Sub Resize(ByVal numCol As Integer, ByVal numRow As Integer)
    Debug.Assert numCol >= k_MIN_MAP_WIDTH And numCol <= k_MAX_MAP_WIDTH
    Debug.Assert numRow >= k_MIN_MAP_HEIGHT And numRow <= k_MAX_MAP_HEIGHT
    
    If GetCol(m_startRoom) > numCol Or GetRow(m_startRoom) > numRow Then
        m_startRoom = 0
    End If
    
    m_mapWidth = numCol
    m_mapHeight = numRow
'    CreateMap numCol, numRow
    FlagMapChange
End Sub

Public Property Let StartRoom(ByVal roomNum As Integer)
    m_startRoom = roomNum
    FlagMapChange
End Property

Public Property Get StartRoom() As Integer
    StartRoom = m_startRoom
End Property

Public Function Height() As Integer
    Height = m_mapHeight
End Function

Public Function Width() As Integer
    Width = m_mapWidth
End Function

Public Function GetRoom(ByVal col As Integer, ByVal row As Integer) As CRoom
    Debug.Assert col >= 0 And col < m_mapWidth
    Debug.Assert row >= 0 And row < m_mapHeight
    If col < 0 Or col >= m_mapWidth Or row < 0 Or row >= m_mapHeight Then
        Set GetRoom = Null
    Else
        Set GetRoom = m_room(col, row)
    End If
    
End Function

Private Function GetCol(ByVal roomNum As Integer) As Integer
    GetCol = roomNum \ m_mapHeight
End Function

Private Function GetRow(ByVal roomNum As Integer) As Integer
    GetRow = roomNum Mod m_mapHeight
End Function

Public Function GetRoomNum(ByVal col As Integer, ByVal row As Integer) As Integer
    GetRoomNum = col * Height + row
End Function

Private Sub Class_Initialize()
    m_hasChanged = False
    CreateMap k_MAX_MAP_WIDTH, k_MAX_MAP_HEIGHT
End Sub

Private Sub CreateMap(ByVal w As Integer, ByVal h As Integer)
    Dim col As Integer
    Dim row As Integer
    Dim newRoom As CRoom
    
    m_startRoom = 0
    m_mapWidth = w
    m_mapHeight = h
    For row = 0 To m_mapHeight - 1
        For col = 0 To m_mapWidth - 1
            Set newRoom = New CRoom
            newRoom.column = col
            newRoom.row = row
            Set m_room(col, row) = newRoom
        Next
    Next
End Sub

Public Sub SerialOut(ByRef proj As CProject)
    Dim prop As PropertyBag
    Dim col As Integer
    Dim row As Integer
    Dim result As Boolean
    Dim id As String
    
'    formProgress.StartProgress "Compressing Level", "Compressing Level Data", m_rows
    Set prop = New PropertyBag
    prop.WriteProperty "Cols", m_mapWidth
    prop.WriteProperty "Rows", m_mapHeight
    prop.WriteProperty "StartRoom", m_startRoom
    result = proj.AddContent("Map", prop)
    Debug.Assert result <> False
    
    For row = 0 To m_mapHeight - 1
        For col = 0 To m_mapWidth - 1
            id = "Cell-" & Trim(col) & "-" & Trim(row)
            m_room(col, row).SerialOut id, proj
        Next
'        formProgress.UpdateProgress row
    Next
    
'    formProgress.EndProgress
End Sub

Public Sub SerialIn(ByRef proj As CProject)
    Dim col As Integer
    Dim numCols As Integer
    Dim row As Integer
    Dim numRows As Integer
    Dim prop As PropertyBag
    Dim result As Boolean
    Dim id As String
    Dim StartRoom As Integer
    
    Set prop = New PropertyBag
    result = proj.GetContent("Map", prop)
    Debug.Assert result <> False
    numCols = prop.ReadProperty("Cols", k_MAX_MAP_WIDTH)
    numRows = prop.ReadProperty("Rows", k_MAX_MAP_HEIGHT)
    StartRoom = prop.ReadProperty("StartRoom", 0)
'    formProgress.StartProgress "Decompressing Level", "Decompressing Level Data", numRows
    CreateMap numCols, numRows
    m_startRoom = StartRoom
    For col = 0 To m_mapWidth - 1
        For row = 0 To m_mapHeight - 1
            id = "Cell-" & Trim(col) & "-" & Trim(row)
            m_room(col, row).SerialIn id, proj
        Next
'        formProgress.UpdateProgress row
    Next
    
'    formProgress.EndProgress
End Sub

Public Sub ExportHeader(ByRef export As CExport, ByVal label As String)
'    export.Comment "Order of exported bytes"
'    export.Comment "floor              byte    0 to 255"
'    export.Comment "background         byte    0 to 255"
'    export.Comment "exit left          byte    0 to 255"
'    export.Comment "exit right         byte    0 to 255"
'    export.Comment "creature           byte    0 to 255"
'    export.Comment "item               byte    0 to 255"
'    export.Comment "feature            word    16-bit bit mask"
'    export.Comment "       " & k_STR_CONST_FEATURE_SAVE_POINT & "    bit " & Trim(k_FEATURE_SAVE_POINT - 1)
'    export.Comment "       " & k_STR_CONST_FEATURE_LADDER & "        bit " & Trim(k_FEATURE_LADDER - 1)
'    export.Comment "       " & k_STR_CONST_FEATURE_BALLOON & "       bit " & Trim(k_FEATURE_BALLOON - 1)
'    export.Comment "       " & k_STR_CONST_FEATURE_WATERFALL & "     bit " & Trim(k_FEATURE_WATERFALL - 1)
'    export.Comment "       " & k_STR_CONST_FEATURE_LARA & "          bit " & Trim(k_FEATURE_LARA - 1)
'    export.Comment "       " & k_STR_CONST_FEATURE_VINE & "          bit " & Trim(k_FEATURE_VINE - 1)
'    export.CRLF 2
'
'    export.ConditionalDefine "ROOM_ENUMS"
'    ' floor equates
'    export.Comment "Floors"
'    export.EQU k_STR_CONST_FLOOR_NONE, k_FLOOR_NONE
'    export.EQU k_STR_CONST_FLOOR_PLAT_LEFT, k_FLOOR_PLAT_LEFT
'    export.EQU k_STR_CONST_FLOOR_PLAT_RIGHT, k_FLOOR_PLAT_RIGHT
'    export.EQU k_STR_CONST_FLOOR_PLAT_BOTH, k_FLOOR_PLAT_BOTH
'    export.EQU k_STR_CONST_FLOOR_PLAT_LEFT_WATER, k_FLOOR_PLAT_LEFT_WATER
'    export.EQU k_STR_CONST_FLOOR_PLAT_RIGHT_WATER, k_FLOOR_PLAT_RIGHT_WATER
'    export.EQU k_STR_CONST_FLOOR_PLAT_BOTH_WATER, k_FLOOR_PLAT_BOTH_WATER
'    export.EQU k_STR_CONST_FLOOR_WATER, k_FLOOR_WATER
'    export.EQU k_STR_CONST_FLOOR_SOLID, k_FLOOR_SOLID
'    export.EQU k_STR_CONST_FLOOR_WALKWAY, k_FLOOR_WALKWAY
'    export.EQU k_STR_CONST_FLOOR_WALKWAY_SINGLE_HOLE, k_FLOOR_WALKWAY_SINGLE_HOLE
'    export.EQU k_STR_CONST_FLOOR_WALKWAY_THREE_HOLES, k_FLOOR_WALKWAY_THREE_HOLES
'    export.EQU k_STR_CONST_FLOOR_RIVER_BED, k_FLOOR_RIVER_BED
'    export.EQU k_STR_CONST_FLOOR_WALKWAY_HOLE_WITH_LADDER, k_FLOOR_WALKWAY_HOLE_WITH_LADDER
'    export.EQU k_STR_CONST_FLOOR_WALKWAY_SIX_HOLES, k_FLOOR_WALKWAY_SIX_HOLES
'    export.CRLF 2
'    ' exit left equates
'    export.Comment "Exits to the left"
'    export.EQU k_STR_CONST_EXIT_LEFT_OPEN, k_EXIT_LEFT_OPEN
'    export.EQU k_STR_CONST_EXIT_LEFT_DARK_ROCK_BLUE, k_EXIT_LEFT_DARK_ROCK_BLUE
'    export.EQU k_STR_CONST_EXIT_LEFT_LIGHT_ROCK_BLACK, k_EXIT_LEFT_LIGHT_ROCK_BLACK
'    export.EQU k_STR_CONST_EXIT_LEFT_DARK_ROCK_BLACK, k_EXIT_LEFT_DARK_ROCK_BLACK
'    export.EQU k_STR_CONST_EXIT_LEFT_LIGHT_ROCK_GREEN, k_EXIT_LEFT_LIGHT_ROCK_GREEN
'    export.EQU k_STR_CONST_EXIT_LEFT_DARK_ROCK_GREEN, k_EXIT_LEFT_DARK_ROCK_GREEN
'    export.EQU k_STR_CONST_EXIT_LEFT_PATTERN_ROCK_BLACK, k_EXIT_LEFT_PATTERN_ROCK_BLACK
'    export.CRLF 2
'    ' exit right equates
'    export.Comment "Exits to the right"
'    export.EQU k_STR_CONST_EXIT_RIGHT_OPEN, k_EXIT_RIGHT_OPEN
'    export.EQU k_STR_CONST_EXIT_RIGHT_DARK_ROCK_BLUE, k_EXIT_RIGHT_DARK_ROCK_BLUE
'    export.EQU k_STR_CONST_EXIT_RIGHT_LIGHT_ROCK_BLACK, k_EXIT_RIGHT_LIGHT_ROCK_BLACK
'    export.EQU k_STR_CONST_EXIT_RIGHT_DARK_ROCK_BLACK, k_EXIT_RIGHT_DARK_ROCK_BLACK
'    export.EQU k_STR_CONST_EXIT_RIGHT_LIGHT_ROCK_GREEN, k_EXIT_RIGHT_LIGHT_ROCK_GREEN
'    export.EQU k_STR_CONST_EXIT_RIGHT_DARK_ROCK_GREEN, k_EXIT_RIGHT_DARK_ROCK_GREEN
'    export.EQU k_STR_CONST_EXIT_RIGHT_PATTERN_ROCK_BLACK, k_EXIT_RIGHT_PATTERN_ROCK_BLACK
'    export.CRLF 2
'
'    ' background equates
''    export.Comment "Backgrounds"
''    export.EQU k_STR_CONST_BACKGROUND_NONE, k_BACKGROUND_NONE
''    export.EQU k_STR_CONST_BACKGROUND_TREES, k_BACKGROUND_TREES
''    export.EQU k_STR_CONST_BACKGROUND_TREE_TOPS, k_BACKGROUND_TREE_TOPS
''    export.EQU k_STR_CONST_BACKGROUND_WATER, k_BACKGROUND_WATER
''    export.EQU k_STR_CONST_BACKGROUND_EARTH, k_BACKGROUND_EARTH
''    export.CRLF 2
'
'    ' feature equates
'    export.Comment "Features"
'    export.EQU k_STR_CONST_FEATURE_SAVE_POINT, k_FEATURE_SAVE_POINT
'    export.EQU k_STR_CONST_FEATURE_LADDER, k_FEATURE_LADDER
'    export.EQU k_STR_CONST_FEATURE_BALLOON, k_FEATURE_BALLOON
'    export.EQU k_STR_CONST_FEATURE_WATERFALL, k_FEATURE_WATERFALL
'    export.EQU k_STR_CONST_FEATURE_LARA, k_FEATURE_LARA
'    export.EQU k_STR_CONST_FEATURE_VINE, k_FEATURE_VINE
'    export.Comment "Features (LSB bit masks)"
'    export.EQU k_STR_CONST_FEATURE_SAVE_POINT_MSK, LShift(1, k_FEATURE_SAVE_POINT - 1)
'    export.EQU k_STR_CONST_FEATURE_LADDER_MSK, LShift(1, k_FEATURE_LADDER - 1)
'    export.EQU k_STR_CONST_FEATURE_BALLOON_MSK, LShift(1, k_FEATURE_BALLOON - 1)
'    export.EQU k_STR_CONST_FEATURE_WATERFALL_MSK, LShift(1, k_FEATURE_WATERFALL - 1)
'    export.EQU k_STR_CONST_FEATURE_LARA_MSK, LShift(1, k_FEATURE_LARA - 1)
'    export.EQU k_STR_CONST_FEATURE_VINE_MSK, LShift(1, k_FEATURE_VINE - 1)
'    export.CRLF 2
'
'    ' creature equates
'    export.Comment "Creatures"
'    export.EQU k_STR_CONST_CREATURE_NONE, k_CREATURE_NONE
'    export.EQU k_STR_CONST_CREATURE_BAT, k_CREATURE_BAT
'    export.EQU k_STR_CONST_CREATURE_CONDOR, k_CREATURE_CONDOR
'    export.EQU k_STR_CONST_CREATURE_EEL, k_CREATURE_EEL
'    export.EQU k_STR_CONST_CREATURE_FROG, k_CREATURE_FROG
'    export.EQU k_STR_CONST_CREATURE_SCORPION, k_CREATURE_SCORPION
'    export.EQU k_STR_CONST_CREATURE_RABID_BAT, k_CREATURE_RABID_BAT
'    export.EQU k_STR_CONST_CREATURE_FIRE_ANT, k_CREATURE_FIRE_ANT
'    export.EQU k_STR_CONST_CREATURE_PIRANHA, k_CREATURE_PIRANHA
'    export.EQU k_STR_CONST_CREATURE_WALk_FROG, k_CREATURE_WALK_FROG
'    export.EQU k_STR_CONST_CREATURE_SNAKE, k_CREATURE_SNAKE
'    export.EQU k_STR_CONST_CREATURE_TREE_SNAKE, k_CREATURE_TREE_SNAKE
'    export.EQU k_STR_CONST_CREATURE_CROCODILE, k_CREATURE_CROCODILE
'    export.CRLF 2
'
'    ' item equates
'    export.Comment "Items"
'    export.EQU k_STR_CONST_ITEM_NONE, k_ITEM_NONE
'    export.EQU k_STR_CONST_ITEM_STONE_RAT, k_ITEM_STONE_RAT
'    export.EQU k_STR_CONST_ITEM_QUICKCLAW_CAT, k_ITEM_QUICKCLAW_CAT
'    export.EQU k_STR_CONST_ITEM_DIAMOND_RING, k_ITEM_DIAMOND_RING
'    export.EQU k_STR_CONST_ITEM_RHONDA_GIRL, k_ITEM_RHONDA_GIRL
'    export.EQU k_STR_CONST_ITEM_GOLD_BAR_LEFT, k_ITEM_GOLD_BAR_LEFT
'    export.EQU k_STR_CONST_ITEM_GOLD_BAR_RIGHT, k_ITEM_GOLD_BAR_RIGHT
'    export.EQU k_STR_CONST_ITEM_HAT, k_ITEM_HAT
'    export.EQU k_STR_CONST_ITEM_LARA, k_ITEM_LARA
'    export.EQU k_STR_CONST_ITEM_LAMP, k_ITEM_LAMP
'    export.EQU k_STR_CONST_ITEM_ROPE, k_ITEM_ROPE
'    export.CRLF 2
'    export.ConditionalEnd
'    export.CRLF 2
'
'    ' export width & height equates
'    export.Comment "Map dimensions"
'    export.EQU "k_" & Trim(UCase(label)) & "_WIDTH", m_mapWidth
'    export.EQU "k_" & Trim(UCase(label)) & "_HEIGHT", m_mapHeight
'    export.CRLF 2
'
'    ' export global declaration
'    export.DeclareGlobal label
End Sub

Public Sub ExportSource(ByRef export As CExport, ByVal label As String)
    Dim col As Integer
    Dim row As Integer
    
'    formProgress.StartProgress "Exporting Level", "Exporting Level Data", m_rows
    export.Text label & "::"
    For col = 0 To m_mapWidth - 1
        For row = 0 To m_mapHeight - 1
            export.DB m_room(col, row).ExportSource, "room " & Trim(col) & ", " & Trim(row)
        Next
        export.CRLF
'        formProgress.UpdateProgress row
    Next
    
'    formProgress.EndProgress
End Sub

Private Sub ExportVCSExits(ByRef export As CVCSImage)
    Dim col As Integer
    Dim row As Integer
    Dim addr As Long
    Dim i As Long
    
    addr = 3174
    For i = 1 To 32
        export.SetByte addr, 0
        addr = addr + 1
    Next
    
End Sub

Public Sub ExportBinary(ByRef export As CVCSImage)
'    ExportVCSExits export
'    ExportVCSRooms export
End Sub

'Public Property Get Serialize() As PropertyBag
'    Dim prop As PropertyBag
'    Dim col As Integer
'    Dim row As Integer
'    Dim cellProp As PropertyBag
'    Dim tempProp As PropertyBag
'
''    formProgress.StartProgress "Compressing Level", "Compressing Level Data", m_rows
'    Set prop = New PropertyBag
'    prop.WriteProperty "Cols", m_mapWidth
'    prop.WriteProperty "Rows", m_mapHeight
'    prop.WriteProperty "StartRoom", m_startRoom
''    prop.WriteProperty "Pic", m_backgroundPicture, Nothing
'    Set cellProp = New PropertyBag
'    Set tempProp = New PropertyBag
'    For row = 0 To m_mapHeight - 1
'        For col = 0 To m_mapWidth - 1
'            cellProp.Contents = tempProp.Contents
'            m_room(col, row).SerializeOut cellProp
'            prop.WriteProperty "Cell-" & Trim(col) & "-" & Trim(row), cellProp.Contents
'        Next
''        formProgress.UpdateProgress row
'    Next
'
'    Set Serialize = prop
''    formProgress.EndProgress
'End Property
'
'
'Public Property Set Serialize(ByRef prop As PropertyBag)
'    Dim col As Integer
'    Dim numCols As Integer
'    Dim row As Integer
'    Dim numRows As Integer
'    Dim cellProp As PropertyBag
'    Dim StartRoom As Integer
'
'    numCols = prop.ReadProperty("Cols", k_MAX_MAP_WIDTH)
'    numRows = prop.ReadProperty("Rows", k_MAX_MAP_HEIGHT)
'    StartRoom = prop.ReadProperty("StartRoom", 0)
''    formProgress.StartProgress "Decompressing Level", "Decompressing Level Data", numRows
''    Set m_backgroundPicture = prop.ReadProperty("Pic")
''    m_initialised = False
'    CreateMap numCols, numRows
'    m_startRoom = StartRoom
'    Set cellProp = New PropertyBag
'    For col = 0 To m_mapWidth - 1
'        For row = 0 To m_mapHeight - 1
'            cellProp.Contents = prop.ReadProperty("Cell-" & Trim(col) & "-" & Trim(row))
'            m_room(col, row).SerializeIn cellProp
'        Next
''        formProgress.UpdateProgress row
'    Next
'
''    formProgress.EndProgress
'End Property
'
'

Public Sub Draw()
End Sub
