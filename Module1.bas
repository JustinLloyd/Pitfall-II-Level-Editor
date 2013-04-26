Attribute VB_Name = "General"
Option Explicit

Public Const k_TITLE = "Pitfall 2: Rumble In the Jungle"

Public Const k_FLIGHT_PATH_LEN = 32
Public Const k_MAX_ANIM_FRAMES = 16

Public Const k_MIN_MAP_WIDTH = 4
Public Const k_MAX_MAP_WIDTH = 64
Public Const k_MIN_MAP_HEIGHT = 4
Public Const k_MAX_MAP_HEIGHT = 64

Public Const k_ROOM_DATA_SIZE As Long = 6&

Public Const k_ROOM_HEIGHT_PIX As Long = 48&
Public Const k_ROOM_WIDTH_PIX As Long = 160&

Public Const k_ROOM_SPACING_X_PIX As Long = k_ROOM_WIDTH_PIX + 2
Public Const k_ROOM_SPACING_Y_PIX As Long = k_ROOM_HEIGHT_PIX + 2

'Public Const k_FLOOR_X = 0&
'Public Const k_FLOOR_Y = 32&

'Public Const k_FEATURE_LADDER_X = 76&
'Public Const k_FEATURE_LADDER_Y = 0&
'Public Const k_FEATURE_SAVE_POINT_X = 40&
'Public Const k_FEATURE_SAVE_POINT_Y = 32&
'Public Const k_FEATURE_WATERFALL_X = 0&
'Public Const k_FEATURE_WATERFALL_Y = 0&
'
'
'Public Const k_ITEM_STONE_RAT_X = 40&
'Public Const k_ITEM_STONE_RAT_Y = 32&
'Public Const k_ITEM_QUICKCLAW_CAT_X = 40&
'Public Const k_ITEM_QUICKCLAW_CAT_Y = 16&
'Public Const k_ITEM_DIAMOND_RING_X = 40&
'Public Const k_ITEM_DIAMOND_RING_Y = 24&
'Public Const k_ITEM_RHONDA_GIRL_X = 40&
'Public Const k_ITEM_RHONDA_GIRL_Y = 16&
'Public Const k_ITEM_HAT_X = 120&
'Public Const k_ITEM_HAT_Y = 32&
'Public Const k_ITEM_LARA_X = 120&
'Public Const k_ITEM_LARA_Y = 32&
'Public Const k_ITEM_GOLD_BAR_LEFT_X = 20&
'Public Const k_ITEM_GOLD_BAR_LEFT_Y = 20&
'Public Const k_ITEM_GOLD_BAR_RIGHT_X = 120&
'Public Const k_ITEM_GOLD_BAR_RIGHT_Y = 20&
'Public Const k_ITEM_LAMP_X = 120&
'Public Const k_ITEM_LAMP_Y = 32&
'Public Const k_ITEM_ROPE_X = 120&
'Public Const k_ITEM_ROPE_Y = 32&
'
'Public Const k_CREATURE_BAT_Y = 8&
'Public Const k_CREATURE_SCORPION_Y = 24&
'Public Const k_CREATURE_CONDOR_Y = 8&
'Public Const k_CREATURE_EEL_Y = 16&
'Public Const k_CREATURE_FIRE_ANT_Y = 32&

Public g_configDB As Database
Public g_testDB As Database
Public g_backgroundType() As CBackground
Public g_hazardType() As CHazard
Public g_itemType() As CItem
Public g_floorType() As CFloor
Public g_groundType() As CGround
Public g_exitType() As CExit
Public g_featureType() As CFeature

Public g_flightPath() As D3DVECTOR2
Public g_projectFilename As String
Public g_map As CMap
Public g_frameCount As Long
Public g_d3dObject As CDirect3D
Public g_d3dDevice As Direct3DDevice8
Public g_d3dX As D3DX8

Public Sub CreateBackgroundTypes()
    Dim rst As Recordset
    Dim newBkg As CBackground
    Dim i As Integer
    
    Set rst = g_configDB.OpenRecordset("SELECT * FROM BackgroundType ORDER BY ID ASC;")
    i = 1
    While Not rst.EOF
        Set newBkg = New CBackground
        Call newBkg.Initialise(rst)
        ReDim Preserve g_backgroundType(1 To i)
        Set g_backgroundType(i) = newBkg
        i = i + 1
        rst.MoveNext
    Wend

    rst.Close
End Sub

Public Sub CreateHazardTypes()
    Dim rst As Recordset
    Dim newHazard As CHazard
    Dim i As Integer
    
    Set rst = g_configDB.OpenRecordset("SELECT * FROM HazardType ORDER BY ID ASC;")
    i = 1
    While Not rst.EOF
        Set newHazard = New CHazard
        Call newHazard.Initialise(rst)
        ReDim Preserve g_hazardType(1 To i)
        Set g_hazardType(i) = newHazard
        i = i + 1
        rst.MoveNext
    Wend

    rst.Close
End Sub

Public Sub CreateItemTypes()
    Dim rst As Recordset
    Dim newItem As CItem
    Dim i As Integer
    
    Set rst = g_configDB.OpenRecordset("SELECT * FROM ItemType ORDER BY ID ASC;")
    i = 1
    While Not rst.EOF
        Set newItem = New CItem
        Call newItem.Initialise(rst)
        ReDim Preserve g_itemType(1 To i)
        Set g_itemType(i) = newItem
        i = i + 1
        rst.MoveNext
    Wend

    rst.Close
End Sub

Public Sub CreateFloorTypes()
    Dim rst As Recordset
    Dim newFloor As CFloor
    Dim i As Integer
    
    Set rst = g_configDB.OpenRecordset("SELECT * FROM FloorType ORDER BY ID ASC;")
    i = 1
    While Not rst.EOF
        Set newFloor = New CFloor
        Call newFloor.Initialise(rst)
        ReDim Preserve g_floorType(1 To i)
        Set g_floorType(i) = newFloor
        i = i + 1
        rst.MoveNext
    Wend

    rst.Close
End Sub

Public Sub CreateGroundTypes()
    Dim rst As Recordset
    Dim newGround As CGround
    Dim i As Integer
    
    Set rst = g_configDB.OpenRecordset("SELECT * FROM GroundType ORDER BY ID ASC;")
    i = 1
    While Not rst.EOF
        Set newGround = New CGround
        Call newGround.Initialise(rst)
        ReDim Preserve g_groundType(1 To i)
        Set g_groundType(i) = newGround
        i = i + 1
        rst.MoveNext
    Wend

    rst.Close
End Sub

Public Sub CreateExitTypes()
    Dim rst As Recordset
    Dim newExit As CExit
    Dim i As Integer
    
    Set rst = g_configDB.OpenRecordset("SELECT * FROM ExitType ORDER BY ID ASC;")
    i = 1
    While Not rst.EOF
        Set newExit = New CExit
        Call newExit.Initialise(rst)
        ReDim Preserve g_exitType(1 To i)
        Set g_exitType(i) = newExit
        i = i + 1
        rst.MoveNext
    Wend

    rst.Close
End Sub

Public Sub CreateFeatureTypes()
    Dim rst As Recordset
    Dim newFeature As CFeature
    Dim i As Integer
    
    Set rst = g_configDB.OpenRecordset("SELECT * FROM FeatureType ORDER BY ID ASC;")
    i = 1
    While Not rst.EOF
        Set newFeature = New CFeature
        Call newFeature.Initialise(rst)
        ReDim Preserve g_featureType(1 To i)
        Set g_featureType(i) = newFeature
        i = i + 1
        rst.MoveNext
    Wend

    rst.Close
End Sub

Public Sub InitFlightPaths()
    Dim rst As Recordset
    Dim i As Integer
    
    Set rst = g_configDB.OpenRecordset("SELECT * FROM FlightPathVertex ORDER BY ID ASC;")
    i = 0
    While Not rst.EOF
        ReDim Preserve g_flightPath(0 To i)
        g_flightPath(i).x = rst!x
        g_flightPath(i).y = rst!y
        i = i + 1
        rst.MoveNext
    Wend

    rst.Close
End Sub


Public Sub CenterForm(ByRef frm As Form)
    frm.Left = (Screen.Width - frm.Width) / 2
    frm.Top = (Screen.Height - frm.Height) / 2
End Sub

Public Function GetDisclaimer() As String
    Dim s As String
    
    s = "Warning: This computer program is protected by copyright law and international treaties. "
    s = s & "Unauthorized reproduction or distribution of this program, or any portion of it, "
    s = s & "may result in severe civil and criminal penalties, and will be prosecuted "
    s = s & "to the maximum extent possible under law."
    GetDisclaimer = s
End Function

Public Function GetCopyright() As String
    GetCopyright = "Copyright © 1999 Otaku No Zoku"
End Function

Public Sub InsizeForm(ByRef frm As Form, ByVal w As Long, ByVal h As Long)
    frm.Width = w + (frm.Width - frm.ScaleX(frm.ScaleWidth, frm.ScaleMode, vbTwips))
    frm.Height = h + (frm.Height - frm.ScaleY(frm.ScaleHeight, frm.ScaleMode, vbTwips))
End Sub

Public Sub UpdateRedrawTime(ByRef status As statusBar, ByVal panelID As String, ByVal redrawTime As Long)
    Dim s As String
    
    s = "  Redraw: " & redrawTime & " ms"
    status.Panels.Item(panelID).Text = s
End Sub

Public Sub Main()
    Dim splashComplete As Boolean

    frmSplash.Show
    DoEvents
    g_frameCount = 0
    Set g_configDB = OpenDatabase(App.path & "\config.mdb")

    Load formMain
    Initialise
    While frmSplash.Timer1.Enabled = True
        DoEvents
    Wend

    Unload frmSplash
    formMain.Show
End Sub

Public Sub createdb()
    Dim i As Integer
    Dim td As TableDef
    Dim prp As Property
    Dim idx As index
    
    Set g_testDB = CreateDatabase(App.path & "\test.mdb", dbLangGeneral)
    Set td = g_testDB.CreateTableDef("Test")
    Set idx = td.CreateIndex("ID")
    idx.Fields = "ID"
    idx.Primary = True
    idx.Unique = True

    td.Fields.Append idx.CreateField("ID", dbLong)
    td.Fields("ID").Attributes = dbAutoIncrField
    td.Fields("ID").AllowZeroLength = False

    td.Fields.Append td.CreateField("Col", dbInteger)
    td.Fields.Append td.CreateField("Row", dbInteger)
    td.Fields.Append td.CreateField("HasChanged", dbBoolean)
    td.Indexes.Append idx

    td.Fields.Refresh
    g_testDB.TableDefs.Append td

    Set idx = Nothing
    Set td = Nothing
    g_testDB.Close
    Set g_testDB = Nothing
    End

End Sub

'Public Sub DrawPlayerSprite(ByVal hdcBuffer As Long, ByVal screenX As Long, ByVal screenY As Long)
'    Dim hdcOld As Long
'    Dim ret As Long
'    Dim hdcPlayer As Long
'
'    If Not CreateRoomDC(hdcBuffer, hdcPlayer) Then Exit Sub
'
'    hdcOld = SelectObject(hdcPlayer, g_picPlayer.m_mask(0))
'    ret = BitBlt(hdcBuffer, screenX, screenY, g_picPlayer.m_width, g_picPlayer.m_height, hdcPlayer, 0&, 0&, SRCAND)
'
'    ret = SelectObject(hdcPlayer, g_picPlayer.m_frame(0))
'    ret = BitBlt(hdcBuffer, screenX, screenY, g_picPlayer.m_width, g_picPlayer.m_height, hdcPlayer, 0&, 0&, SRCPAINT)
'    ret = SelectObject(hdcPlayer, hdcOld)
'
'    ReleaseRoomDC hdcPlayer
'End Sub
'
'Private Sub DrawRoomCreature(ByVal hdcRoom As Long, ByRef room As CRoom, ByVal frameCount As Long)
'    Dim animFrame As Integer
'    Dim hdcCreature As Long
'    Dim hdcOld As Long
'    Dim ret As Long
'    Dim dx As Long
'    Dim dy As Long
'    Dim flightPathPos As Integer
'
'    If room.Creature = k_CREATURE_NONE Then
'    ElseIf room.Creature = k_CREATURE_BAT Then
'        animFrame = frameCount Mod 2&
'        dx = frameCount Mod 160&
'        flightPathPos = frameCount Mod k_FLIGHT_PATH_LEN
'        dy = k_CREATURE_BAT_Y + g_flightPath(flightPathPos)
'    ElseIf room.Creature = k_CREATURE_SCORPION Then
'        animFrame = frameCount Mod 2&
'        dx = frameCount * 2& Mod 160&
'        dy = k_CREATURE_SCORPION_Y
'    ElseIf room.Creature = k_CREATURE_EEL Then
'        animFrame = frameCount Mod 2&
'        If Int(Rnd * 100) > 50 Then animFrame = animFrame + 2
'        dx = frameCount * 2& Mod 160&
'        dy = k_CREATURE_EEL_Y
'    ElseIf room.Creature = k_CREATURE_CONDOR Then
'        animFrame = frameCount Mod 2&
'        dx = 160 - (frameCount Mod 160&)
'        flightPathPos = (frameCount \ 2&) Mod k_FLIGHT_PATH_LEN
'        dy = k_CREATURE_CONDOR_Y + g_flightPath(flightPathPos)
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
'
'    If room.Creature <> k_CREATURE_NONE Then
'        If Not CreateRoomDC(hdcRoom, hdcCreature) Then Exit Sub
'        With g_picCreature(room.Creature)
'            hdcOld = SelectObject(hdcCreature, .m_mask(animFrame))
'            ret = BitBlt(hdcRoom, dx, dy, .m_width, .m_height, hdcCreature, 0&, 0&, SRCAND)
'
'            ret = SelectObject(hdcCreature, .m_frame(animFrame))
'            ret = BitBlt(hdcRoom, dx, dy, .m_width, .m_height, hdcCreature, 0&, 0&, SRCPAINT)
'        End With
'
'        ret = SelectObject(hdcCreature, hdcOld)
'        ReleaseRoomDC hdcCreature
'    End If
'
'End Sub


Private Sub Initialise()
    Set g_d3dObject = New CDirect3D
    If Not g_d3dObject.Init(formMain.picView.hwnd) Then
        End
    End If

    Set g_d3dDevice = g_d3dObject.GetDirect3DDevice
    Set g_d3dX = g_d3dObject.GetDirect3DX
    CreateBackgroundTypes
    CreateHazardTypes
    CreateItemTypes
    CreateFloorTypes
    CreateGroundTypes
    CreateExitTypes
    CreateFeatureTypes
    InitFlightPaths
End Sub

Private Sub Requiem()
    g_configDB.Close
End Sub

Public Sub ReportError(ByVal s As String)
    MsgBox s, vbOK, "Error"
End Sub
