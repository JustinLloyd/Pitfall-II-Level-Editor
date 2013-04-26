VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form formMain 
   Caption         =   "Pitfall 2 : Rumble in the Jungle -- Level Editor"
   ClientHeight    =   9405
   ClientLeft      =   5910
   ClientTop       =   3825
   ClientWidth     =   9525
   LinkTopic       =   "Form1"
   ScaleHeight     =   9405
   ScaleWidth      =   9525
   Begin VB.PictureBox picView 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   120
      ScaleHeight     =   2025
      ScaleWidth      =   2625
      TabIndex        =   21
      Top             =   960
      Width           =   2655
   End
   Begin ComCtl3.CoolBar CoolBar 
      Height          =   690
      Left            =   240
      TabIndex        =   19
      Top             =   120
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   1217
      BandCount       =   1
      _CBWidth        =   9015
      _CBHeight       =   690
      _Version        =   "6.7.8862"
      Child1          =   "toolbarMain"
      MinWidth1       =   3195
      MinHeight1      =   630
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Begin MSComctlLib.Toolbar toolbarMain 
         Height          =   630
         Left            =   30
         TabIndex        =   20
         Top             =   30
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   1111
         ButtonWidth     =   1217
         ButtonHeight    =   953
         AllowCustomize  =   0   'False
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   12
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "New"
               Key             =   "New"
               Object.ToolTipText     =   "New Project"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Open"
               Key             =   "Open"
               Object.ToolTipText     =   "Open Project"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Caption         =   "Save"
               Key             =   "Save"
               Object.ToolTipText     =   "Save Project"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Cut"
               Key             =   "Cut"
               Object.ToolTipText     =   "Cut"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Copy"
               Key             =   "Copy"
               Object.ToolTipText     =   "Copy"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Paste"
               Key             =   "Paste"
               Object.ToolTipText     =   "Paste"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Undo"
               Key             =   "Undo"
               Object.ToolTipText     =   "Undo the last action"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Redo"
               Key             =   "Redo"
               Object.ToolTipText     =   "Redo the last action"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Animate"
               Key             =   "Animate"
               Object.ToolTipText     =   "Animate the map"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Play"
               Key             =   "Play"
               Object.ToolTipText     =   "Play the level"
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
   End
   Begin MSComctlLib.ImageList imagelistToolbar 
      Left            =   1080
      Top             =   6120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0000
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0112
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0224
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0336
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":049A
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":05FA
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":076A
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":087C
            Key             =   "Redo"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":098E
            Key             =   "Animate"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0AEE
            Key             =   "Pause"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0C4E
            Key             =   "Play"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar statusBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   17
      Top             =   9105
      Width           =   9525
      _ExtentX        =   16801
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
            Key             =   "StartRoomInfo"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
            Key             =   "MapInfo"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "MemoryInfo"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "Redraw"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "Level Options"
      Height          =   735
      Left            =   4200
      TabIndex        =   16
      Top             =   7080
      Width           =   5265
      Begin VB.CheckBox checkStartRoom 
         Caption         =   "Start Room"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   "Mark current room as starting room"
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.HScrollBar hscrollMap 
      Height          =   255
      Left            =   0
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5520
      Width           =   3015
   End
   Begin VB.VScrollBar vscrollMap 
      Height          =   2655
      Left            =   3000
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   960
      Width           =   255
   End
   Begin VB.Timer timerAnimation 
      Interval        =   100
      Left            =   2640
      Top             =   4680
   End
   Begin VB.Frame frameGBOptions 
      Caption         =   "Room Options"
      Height          =   6135
      Left            =   6360
      TabIndex        =   0
      Top             =   840
      Width           =   3105
      Begin VB.ComboBox comboFeature 
         Height          =   315
         ItemData        =   "Main.frx":1190
         Left            =   120
         List            =   "Main.frx":1192
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   2760
         Width           =   2895
      End
      Begin VB.ComboBox comboFloor 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   2040
         Width           =   2895
      End
      Begin VB.ComboBox comboExitRight 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   5640
         Width           =   2895
      End
      Begin VB.ComboBox comboBackground 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   600
         Width           =   2895
      End
      Begin VB.ComboBox comboHazard 
         Height          =   315
         ItemData        =   "Main.frx":1194
         Left            =   120
         List            =   "Main.frx":1196
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   3480
         Width           =   2895
      End
      Begin VB.ComboBox comboGround 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1320
         Width           =   2895
      End
      Begin VB.ComboBox comboItem 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   4200
         Width           =   2895
      End
      Begin VB.ComboBox comboExitLeft 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   4920
         Width           =   2895
      End
      Begin VB.Label Label7 
         Caption         =   "Floor"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Label6 
         Caption         =   "Right Exit"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   5400
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Background"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Ground"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Feature"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Hazard"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "Item"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   3960
         Width           =   1935
      End
      Begin VB.Label Label12 
         Caption         =   "Left Exit"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   4680
         Width           =   1935
      End
   End
   Begin VB.Shape shapeHighlight 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      Height          =   1695
      Left            =   240
      Top             =   3240
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Menu menuFile 
      Caption         =   "&File"
      Begin VB.Menu menuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu menuOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu menuEmpty0 
         Caption         =   "-"
      End
      Begin VB.Menu menuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu menuSaveAs 
         Caption         =   "Sa&ve As..."
      End
      Begin VB.Menu menuEmpty2 
         Caption         =   "-"
      End
      Begin VB.Menu menuExportToGameboy 
         Caption         =   "Export To &Gameboy"
         Shortcut        =   ^G
      End
      Begin VB.Menu menuExportToVCS 
         Caption         =   "Export to &Atari VCS"
         Shortcut        =   ^A
      End
      Begin VB.Menu menuMRU 
         Caption         =   "MRU"
         Index           =   0
      End
      Begin VB.Menu menuEmpty1 
         Caption         =   "-"
      End
      Begin VB.Menu menuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu menuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu menuUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu menuRedo 
         Caption         =   "&Redo"
      End
      Begin VB.Menu menuEmpty4 
         Caption         =   "-"
      End
      Begin VB.Menu menuCut 
         Caption         =   "Cu&t Room"
         Shortcut        =   ^X
      End
      Begin VB.Menu menuCopy 
         Caption         =   "C&opy Room"
         Shortcut        =   ^C
      End
      Begin VB.Menu menuPaste 
         Caption         =   "&Paste Room"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu menuProject 
      Caption         =   "&Project"
      Begin VB.Menu menuProjectSetSize 
         Caption         =   "Prop&erties..."
      End
   End
   Begin VB.Menu menuRoom 
      Caption         =   "&Room"
      Begin VB.Menu menuResetRoom 
         Caption         =   "&Reset Room"
         Shortcut        =   {DEL}
      End
   End
   Begin VB.Menu menuTricks 
      Caption         =   "&Tricks"
      Begin VB.Menu menuLoadBitmaps 
         Caption         =   "&Load Bitmaps"
         Shortcut        =   ^L
      End
      Begin VB.Menu menuTricksCreatures 
         Caption         =   "&Creatures"
         Enabled         =   0   'False
      End
      Begin VB.Menu menuTricksItems 
         Caption         =   "&Items"
         Enabled         =   0   'False
      End
      Begin VB.Menu menuTricksBackground 
         Caption         =   "&Background"
         Enabled         =   0   'False
      End
      Begin VB.Menu menuTricksFloor 
         Caption         =   "&Floor"
         Enabled         =   0   'False
      End
      Begin VB.Menu menuTricksRightExit 
         Caption         =   "&Right Exit"
         Enabled         =   0   'False
      End
      Begin VB.Menu menuTricksLeftExit 
         Caption         =   "&Left Exit"
         Enabled         =   0   'False
      End
      Begin VB.Menu menuTricksFeature 
         Caption         =   "&Feature"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu menuHelp 
      Caption         =   "&Help"
      Begin VB.Menu menuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "formMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' TO DO
'=======
' Add export properties (default export directory/filename, etc)
' Save export properties in project file
' Export version info of program in export file
' Add "Link To" to room properties
' Add multiple undo/redo buffer
' Add some icons to the menus
' Position components on screen (x & y)
' Export path data for condor flight
' Export path data for bat flight

' PITFALL HARRY
' =============
' add rope climbing

' FEATURES
' ========
' breezy updraft
' gas bubbles
' balloon
' dunking turtles
' floating logs
' floating logs that tip you off (balancing act)
' animating leaves for the breezy updraft
' swinging vine
' gas bubbles
' swinging vine
' password point

' BACKGROUNDS
' ===========
' add lava

' FLOORS
' ======
' Add lava

' EXITS
' =====
' add lava rock face
' opening closing hole floor type

' HAZARDS
' =======
' add rolling boulders
' Add rabid bat (creature)
' Add hopping ground frog (creature)
' Add piranha (creature)
' rolling logs
' crocodiles
' make the hole frog hop around


Option Explicit

Private Const k_WINDOW_TITLE = k_TITLE & " -- Level Editor"
Private Const edKeyBackground = vbKey1   ' keycode used to switch background types
Private Const edKeyGround = vbKey2        ' keycode used to switch floor types
Private Const edKeyFloor = vbKey3        ' keycode used to switch floor types
Private Const edKeyFeature = vbKey4      ' keycode used to switch feature types
Private Const edKeyHazard = vbKey5     ' keycode used to switch Hazard types
Private Const edKeyItem = vbKey6         ' keycode used to switch item types
Private Const edKeyLeftExit = vbKey7    ' keycode used to switch left exit types
Private Const edKeyRightExit = vbKey8   ' keycode used to switch right exit types
Private Const edKeyRoomUp = vbKeyUp
Private Const edKeyRoomDown = vbKeyDown
Private Const edKeyRoomLeft = vbKeyLeft
Private Const edKeyRoomRight = vbKeyRight

Private Const vbmaPause = False
Private Const vbmaPlay = True

'Private Type CUSTOMVERTEX
'    x As Single 'x in screen space.
'    y As Single 'y in screen space.
'    z As Single 'normalized z.
'    rhw As Single 'normalized z rhw.
'    color As Long 'vertex color.
'End Type

Private m_updatingProperties As Boolean
Private m_worldCol As Integer
Private m_worldRow As Integer
Private m_currentRoomCol As Integer
Private m_currentRoomRow As Integer
Private m_currentRoom As CRoom
Private m_currentPicture As PictureBox
Private m_copyRoom As CRoom
Private m_mruList As CMRUList
Private m_animateMap As Boolean
Private m_viewWidth As Integer
Private m_viewHeight As Integer

Private Sub checkStartRoom_Click()
    If Not m_updatingProperties Then
        g_map.StartRoom = g_map.GetRoomNum(m_currentRoomCol, m_currentRoomRow)
    End If
End Sub

Private Sub PreviousFeature()
    If comboFeature.ListIndex - 1 >= 0 Then
        comboFeature.ListIndex = comboFeature.ListIndex - 1
    Else
        comboFeature.ListIndex = comboFeature.ListCount - 1
    End If

End Sub

Private Sub NextFeature()
    If comboFeature.ListIndex + 1 < comboFeature.ListCount Then
        comboFeature.ListIndex = comboFeature.ListIndex + 1
    Else
        comboFeature.ListIndex = 0
    End If

End Sub

Private Sub comboFeature_Click()
    If Not m_updatingProperties Then
        Call Action_ModifyFeature(comboFeature.ListIndex)
    End If
    
End Sub

Private Sub comboFloor_Click()
    If Not m_updatingProperties Then
        Call Action_ModifyFloor(comboFloor.ListIndex)
    End If
    
End Sub

Private Sub comboGround_Click()
    If Not m_updatingProperties Then
        Call Action_ModifyGround(comboGround.ListIndex)
    End If
    
End Sub

Private Sub picView_KeyUp(KeyCode As Integer, Shift As Integer)
'    Debug.Print KeyCode
    If Shift Then
        Select Case KeyCode
            Case edKeyBackground
                Call PreviousBackground
            Case edKeyGround
                Call PreviousGround
            Case edKeyFloor
                Call PreviousFloor
            Case edKeyFeature
                Call PreviousFeature
            Case edKeyHazard
                Call PreviousHazard
            Case edKeyItem
                Call PreviousItem
            Case edKeyLeftExit
                Call PreviousExitLeft
            Case edKeyRightExit
                Call PreviousExitRight
        End Select
    Else
        Select Case KeyCode
            Case edKeyBackground
                Call NextBackground
            Case edKeyGround
                Call NextGround
            Case edKeyFloor
                Call NextFloor
            Case edKeyFeature
                Call NextFeature
            Case edKeyHazard
                Call NextHazard
            Case edKeyItem
                Call NextItem
            Case edKeyLeftExit
                Call NextExitLeft
            Case edKeyRightExit
                Call NextExitRight
            Case edKeyRoomUp
                picView.SetFocus
                Call RoomSelectionUp
            Case edKeyRoomDown
                picView.SetFocus
                Call RoomSelectionDown
            Case edKeyRoomLeft
                picView.SetFocus
                Call RoomSelectionLeft
            Case edKeyRoomRight
                picView.SetFocus
                Call RoomSelectionRight
        End Select
    End If
    
End Sub

Private Sub SetWindowTitle()
    Dim projName As String

    If g_projectFilename = "" Then
        projName = "Untitled"
    Else
        projName = g_projectFilename
    End If

    Me.Caption = k_WINDOW_TITLE & " (" & Trim(projName) & ")"
End Sub

Private Sub PreviousBackground()
    If comboBackground.ListIndex - 1 >= 0 Then
        comboBackground.ListIndex = comboBackground.ListIndex - 1
    Else
        comboBackground.ListIndex = comboBackground.ListCount - 1
    End If
End Sub

Private Sub NextBackground()
    If comboBackground.ListIndex + 1 < comboBackground.ListCount Then
        comboBackground.ListIndex = comboBackground.ListIndex + 1
    Else
        comboBackground.ListIndex = 0
    End If
End Sub

Private Sub comboBackground_Click()
    If Not m_updatingProperties Then
        Action_ModifyBackground comboBackground.ListIndex
    End If
End Sub

Private Sub Action_ModifyBackground(ByVal backgroundNum As Integer)
    m_currentRoom.Background = backgroundNum
End Sub

Private Sub PreviousHazard()
    If comboHazard.ListIndex - 1 >= 0 Then
        comboHazard.ListIndex = comboHazard.ListIndex - 1
    Else
        comboHazard.ListIndex = comboHazard.ListCount - 1
    End If

End Sub

Private Sub NextHazard()
    If comboHazard.ListIndex + 1 < comboHazard.ListCount Then
        comboHazard.ListIndex = comboHazard.ListIndex + 1
    Else
        comboHazard.ListIndex = 0
    End If

End Sub

Private Sub comboHazard_Click()
    If Not m_updatingProperties Then
        Action_ModifyHazard comboHazard.ListIndex
    End If
End Sub

Private Sub Action_ModifyHazard(ByVal HazardNum As Integer)
    m_currentRoom.Hazard = HazardNum
End Sub

Private Sub PreviousExitLeft()
    comboExitLeft.ListIndex = IncWrap(comboExitLeft.ListIndex, comboExitLeft.Count - 1, 0)
'    If comboExitLeft.ListIndex - 1 >= 0 Then
'        comboExitLeft.ListIndex = comboExitLeft.ListIndex - 1
'    Else
'        comboExitLeft.ListIndex = comboExitLeft.ListCount - 1
'    End If

End Sub

Private Sub NextExitLeft()
    If comboExitLeft.ListIndex + 1 < comboExitLeft.ListCount Then
        comboExitLeft.ListIndex = comboExitLeft.ListIndex + 1
    Else
        comboExitLeft.ListIndex = 0
    End If

End Sub

Private Sub comboExitLeft_Click()
    If Not m_updatingProperties Then
        Action_ModifyExitLeft comboExitLeft.ListIndex
    End If
End Sub

Private Sub Action_ModifyExitLeft(ByVal exitLeftNum As Integer)
    m_currentRoom.ExitLeft = exitLeftNum
End Sub

Private Sub PreviousExitRight()
    If comboExitRight.ListIndex - 1 >= 0 Then
        comboExitRight.ListIndex = comboExitRight.ListIndex - 1
    Else
        comboExitRight.ListIndex = comboExitRight.ListCount - 1
    End If

End Sub

Private Sub NextExitRight()
    If comboExitRight.ListIndex + 1 < comboExitRight.ListCount Then
        comboExitRight.ListIndex = comboExitRight.ListIndex + 1
    Else
        comboExitRight.ListIndex = 0
    End If

End Sub

Private Sub comboExitRight_Click()
    If Not m_updatingProperties Then
        Action_ModifyExitRight comboExitRight.ListIndex
    End If
End Sub

Private Sub Action_ModifyExitRight(ByVal exitRightNum As Integer)
    m_currentRoom.ExitRight = exitRightNum
End Sub

Private Sub PreviousFloor()
    If comboFloor.ListIndex - 1 >= 0 Then
        comboFloor.ListIndex = comboFloor.ListIndex - 1
    Else
        comboFloor.ListIndex = comboFloor.ListCount - 1
    End If

End Sub

Private Sub NextFloor()
    If comboFloor.ListIndex + 1 < comboFloor.ListCount Then
        comboFloor.ListIndex = comboFloor.ListIndex + 1
    Else
        comboFloor.ListIndex = 0
    End If

End Sub

Private Sub Action_ModifyFloor(ByVal floorNum As Integer)
    m_currentRoom.Floor = floorNum
End Sub

Private Sub PreviousGround()
    If comboGround.ListIndex - 1 >= 0 Then
        comboGround.ListIndex = comboGround.ListIndex - 1
    Else
        comboGround.ListIndex = comboGround.ListCount - 1
    End If

End Sub

Private Sub NextGround()
    If comboGround.ListIndex + 1 < comboGround.ListCount Then
        comboGround.ListIndex = comboGround.ListIndex + 1
    Else
        comboGround.ListIndex = 0
    End If

End Sub

Private Sub Action_ModifyGround(ByVal groundNum As Integer)
    m_currentRoom.Ground = groundNum
End Sub

Private Sub PreviousItem()
    If comboItem.ListIndex - 1 >= 0 Then
        comboItem.ListIndex = comboItem.ListIndex - 1
    Else
        comboItem.ListIndex = comboItem.ListCount - 1
    End If

End Sub

Private Sub NextItem()
    If comboItem.ListIndex + 1 < comboItem.ListCount Then
        comboItem.ListIndex = comboItem.ListIndex + 1
    Else
        comboItem.ListIndex = 0
    End If

End Sub

Private Sub comboItem_Click()
    If Not m_updatingProperties Then
        Action_ModifyItem comboItem.ListIndex
    End If
End Sub

Private Sub Action_ModifyItem(ByVal itemNum As Integer)
    m_currentRoom.Item = itemNum
End Sub

Private Sub Action_ModifyFeature(ByVal featureNum As Integer)
    m_currentRoom.Feature = featureNum
End Sub

Private Sub Form_Load()
    Initialise
End Sub

Private Sub Initialise()
    m_updatingProperties = True
    m_animateMap = vbmaPlay
    Set m_mruList = New CMRUList
    m_mruList.Init menuMRU
    InitToolbars
    InitFormComponents
    DisableProjectCommands
    DisablePaste
    InitOptions
    CreateNewLevel
    Me.KeyPreview = True
    m_updatingProperties = False
End Sub

Private Sub InitFormComponents()
    Width = 12000
    picView.BorderStyle = vbBSNone
    CenterForm Me
End Sub

Private Sub PositionFormComponents()
    Dim y As Single
    
    If Me.WindowState = vbMinimized Then Exit Sub
    
    ScaleMode = vbTwips
    CoolBar.Top = 0
    CoolBar.Left = 0
    CoolBar.Width = ScaleWidth
    CoolBar.Height = 460

    picView.Left = 0
    picView.Top = CoolBar.Height
    picView.Width = ScaleWidth - frameGBOptions.Width - vscrollMap.Width - 120
    picView.Height = ScaleHeight - CoolBar.Height - statusBar.Height - hscrollMap.Height
    
    ScaleMode = vbPixels
    m_viewWidth = picView.Width / k_ROOM_SPACING_X_PIX + 0.5
    m_viewHeight = picView.Height / k_ROOM_SPACING_Y_PIX + 0.5
    Debug.Print m_viewWidth, m_viewHeight
    ScaleMode = vbTwips
    
    vscrollMap.Top = CoolBar.Height
    vscrollMap.Left = picView.Left + picView.Width
    vscrollMap.Height = picView.Height
    hscrollMap.Left = picView.Left
    hscrollMap.Top = picView.Top + picView.Height
    hscrollMap.Width = picView.Width
    
    frameGBOptions.Top = picView.Top
    frameGBOptions.Left = vscrollMap.Left + vscrollMap.Width + 100
    
    Frame2.Left = frameGBOptions.Left
    Frame2.Top = frameGBOptions.Top + frameGBOptions.Height
    
    If Not g_map Is Nothing Then
        Dim maxVal As Integer
        
        maxVal = g_map.Height - m_viewHeight
        vscrollMap.LargeChange = m_viewHeight - 1
        vscrollMap.value = IntMin(maxVal, vscrollMap.value)
        vscrollMap.Max = maxVal

        maxVal = g_map.Width - m_viewWidth
        hscrollMap.LargeChange = m_viewWidth - 1
        hscrollMap.value = IntMin(maxVal, hscrollMap.value)
        hscrollMap.Max = maxVal
    End If
    
End Sub

Private Sub DisableProjectCommands()
    menuSave.Enabled = False
    menuSaveAs.Enabled = False
    menuExportToGameboy.Enabled = False
    menuExportToVCS.Enabled = False
    toolbarMain.Buttons.Item("Save").Enabled = False
End Sub

Private Sub EnableProjectCommands()
    menuSave.Enabled = True
    menuSaveAs.Enabled = True
    menuExportToGameboy.Enabled = True
    menuExportToVCS.Enabled = True
    toolbarMain.Buttons.Item("Save").Enabled = True
End Sub

Private Function VerifyExitWithoutSaving() As Boolean
    Dim s As String

    s = "The current level has been altered but not saved." & vbCrLf
    s = s + "Are you sure you want to exit?"
    If MsgBox(s, vbYesNo, "Warning: Save Level") = vbYes Then
        VerifyExitWithoutSaving = True
    Else
        VerifyExitWithoutSaving = False
    End If
End Function

Private Function VerifyNewWithoutSaving() As Boolean
    Dim s As String

    s = "The current level has been altered but not saved." & vbCrLf
    s = s + "Are you sure you want to create a new level?"
    If MsgBox(s, vbYesNo, "Warning: Save Level") = vbYes Then
        VerifyNewWithoutSaving = True
    Else
        VerifyNewWithoutSaving = False
    End If
End Function

Private Function VerifyOpenWithoutSaving() As Boolean
    Dim s As String

    s = "The current level has been altered but not saved." & vbCrLf
    s = s + "Are you sure you want to open another level?"
    If MsgBox(s, vbYesNo, "Warning: Save Level") = vbYes Then
        VerifyOpenWithoutSaving = True
    Else
        VerifyOpenWithoutSaving = False
    End If
End Function

Private Sub Form_Resize()
    Dim vp As D3DVIEWPORT8
    
    PositionFormComponents
    ScaleMode = vbPixels
    vp.x = 0
    vp.y = 0
    vp.Width = picView.Width
    vp.Height = picView.Height
    If Not g_d3dDevice Is Nothing Then
'        Call g_d3dDevice.SetViewport(vp)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' if current project has changed
    If g_map.HasChanged Then
        ' if user didn't really want to exit
        If Not VerifyExitWithoutSaving Then
            ' cancel exit
            Cancel = True
            ' exit function
            Exit Sub
        End If
    End If

    ' clear up current project
    Set g_map = Nothing
    frmGame.PauseGame
    DoEvents
    Sleep 50
    DoEvents
    Unload frmGame
End Sub

Private Sub SetScrollBars()
    vscrollMap.SmallChange = 1
    vscrollMap.LargeChange = m_viewHeight - 1
    vscrollMap.Min = 0
    vscrollMap.Max = g_map.Height - m_viewHeight
    vscrollMap.value = 0

    hscrollMap.SmallChange = 1
    hscrollMap.LargeChange = m_viewWidth - 1
    hscrollMap.Min = 0
    hscrollMap.Max = g_map.Width - m_viewWidth
    hscrollMap.value = 0
End Sub

Private Sub InitMapPosition()
    m_worldCol = 0
    m_worldRow = 0
    m_currentRoomCol = 0
    m_currentRoomRow = 0
    SetScrollBars
    Set m_currentRoom = g_map.GetRoom(0, 0)
    UpdateRoomProperties m_currentRoom
    UpdateMapInfo
End Sub

Private Sub InitOptions()
    PopulateBackgroundTypeListBox
    PopulateHazardTypeListBox
    PopulateItemListBox
    PopulateFloorTypeListBox
    PopulateGroundTypeListBox
    PopulateExitTypeListBox
    PopulateFeatureTypeListBox
End Sub

Private Sub ComboReset(ByRef cb As ComboBox)
    cb.Clear
    cb.AddItem "None", 0
    cb.ItemData(cb.NewIndex) = 0
    cb.ListIndex = 0
End Sub

Private Sub ComboAdd(ByRef cb As ComboBox, ByVal s As String, ByVal id As Long)
    cb.AddItem s
    cb.ItemData(cb.NewIndex) = id
End Sub

Private Sub PopulateListBox(ByRef cb As ComboBox, ByVal tableName As String)
    Dim sql As String
    Dim rst As Recordset
    
    Call ComboReset(cb)
    Set rst = g_configDB.OpenRecordset("SELECT * FROM " & Trim(tableName) & " ORDER BY ID ASC;")
    While Not rst.EOF
        Call ComboAdd(cb, rst!ListBoxName, rst!id)
        rst.MoveNext
    Wend
    
    rst.Close
End Sub

Private Sub PopulateBackgroundTypeListBox()
    Call PopulateListBox(comboBackground, "BackgroundType")
End Sub

Private Sub PopulateHazardTypeListBox()
    Call PopulateListBox(comboHazard, "HazardType")
End Sub

Private Sub PopulateItemListBox()
    Call PopulateListBox(comboItem, "ItemType")
End Sub

Private Sub PopulateFloorTypeListBox()
    Call PopulateListBox(comboFloor, "FloorType")
End Sub

Private Sub PopulateGroundTypeListBox()
    Dim rst As Recordset
    
    Call ComboReset(comboGround)
    Set rst = g_configDB.OpenRecordset("SELECT * FROM GroundType ORDER BY ID ASC;")
    While Not rst.EOF
        Call ComboAdd(comboGround, rst!ListBoxName, rst!id)
        rst.MoveNext
    Wend
    
    rst.Close
End Sub

Private Sub PopulateExitTypeListBox()
    Call PopulateListBox(comboExitLeft, "ExitType")
    Call PopulateListBox(comboExitRight, "ExitType")
End Sub

Private Sub PopulateFeatureTypeListBox()
    Call PopulateListBox(comboFeature, "FeatureType")
End Sub

Private Sub hscrollMap_Change()
    m_worldCol = hscrollMap.value
End Sub

Private Sub hscrollMap_Scroll()
    m_worldCol = hscrollMap.value
End Sub

Private Sub menuCopy_Click()
    Action_CopyRoom
End Sub

Private Sub menuCut_Click()
    Action_CutRoom
End Sub

Private Sub menuExit_Click()
    Unload Me
End Sub

Private Sub menuExportToGameboy_Click()
    ExportLevel
End Sub

Private Sub menuExportToVCS_Click()
    ExportLevel
End Sub

Private Sub menuHelpAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub menuLoadBitmaps_Click()
'    LoadBitmaps
End Sub

Private Sub menuMRU_Click(index As Integer)
    VerifiedLoadLevelFromMRU index
End Sub

Private Sub menuNew_Click()
    VerifiedNewLevel
End Sub

Private Sub VerifiedNewLevel()
    ' if current project has changed
    If g_map.HasChanged Then
            ' if user didn't really want to exit
            If Not VerifyNewWithoutSaving Then
                ' exit function
                Exit Sub
            End If
    End If

    CreateNewLevel
End Sub
Private Sub menuOpen_Click()
    VerifiedLoadLevel
End Sub

Private Sub menuPaste_Click()
    Action_PasteRoom
End Sub

Private Sub menuProjectSetSize_Click()
    formProjectProperties.Show vbModal
    InitMapPosition
    UpdateMapInfo
End Sub

Private Sub menuResetRoom_Click()
    Action_ResetRoom
End Sub

Private Sub menuSave_Click()
    SaveLevel
End Sub

Private Sub SaveLevelToFile(ByVal Filename As String)
    Dim proj As CProject

    Set proj = New CProject
    g_map.SerialOut proj
    If Not proj.WriteContents(Filename) Then
        MsgBox "Failed to save level", , "File Error"
    Else
        g_map.ClearChange
    End If

End Sub

Private Function LoadLevelFromFile(ByVal Filename As String) As Boolean
    Dim proj As CProject

    Set proj = New CProject
    If Not proj.ReadContents(Filename) Then
        MsgBox "Failed to load level", , "File Error"
        LoadLevelFromFile = False
    Else
        Set g_map = New CMap
        g_map.SerialIn proj
        g_map.ClearChange
        LoadLevelFromFile = True
    End If

End Function

Private Sub menuSaveAs_Click()
    SaveLevelAs
End Sub

Private Sub menuTricksHazards_Click()
    formCreatureEditor.Show vbModal
End Sub

Private Sub UpdateRoomNumber()
    Dim s As String

    s = "Room #" & Trim(m_currentRoomCol * g_map.Height + m_currentRoomRow)
    s = s & "  (" & Trim(m_currentRoomCol) & ", " & Trim(m_currentRoomRow)
    s = s & ") of " & Trim(g_map.Width * g_map.Height)
    s = s & "  (" & Trim(g_map.Width) & ", " & Trim(g_map.Height) & ")"
    statusBar.Panels.Item("MapInfo").Text = s
End Sub

Private Sub UpdateStartRoom()
    Dim s As String

    s = "Start Room #" & Trim(g_map.StartRoom) & "  (" & Trim(g_map.StartRoom \ g_map.Height) & ", " & Trim(g_map.StartRoom Mod g_map.Height) & ")"
    statusBar.Panels.Item("StartRoomInfo").Text = s
'    labelStartRoom.Caption = s
End Sub

Private Sub UpdateRoomProperties(ByRef room As CRoom)
    m_updatingProperties = True
    If g_map.GetRoomNum(m_currentRoomCol, m_currentRoomRow) = g_map.StartRoom Then
        checkStartRoom.value = 1
    Else
        checkStartRoom.value = 0
    End If

    On Error Resume Next
    comboGround.ListIndex = room.Ground
    comboFloor.ListIndex = room.Floor
    comboExitLeft.ListIndex = room.ExitLeft
    comboExitRight.ListIndex = room.ExitRight
    comboBackground.ListIndex = room.Background
    comboItem.ListIndex = room.Item
    comboHazard.ListIndex = room.Hazard
    comboFeature.ListIndex = room.Feature
    m_updatingProperties = False
End Sub

Private Sub HighlightRoom()
    Dim lineList(0 To 4) As TLVERTEX
    Dim sizeOfVertex As Long
    Dim pos As D3DVECTOR2
    Dim vb As Direct3DVertexBuffer8
    Dim r As Integer
    Dim g As Integer
    Dim b As Integer
    Static flashOn As Boolean
    
    sizeOfVertex = Len(lineList(0))
    
    Set vb = g_d3dDevice.CreateVertexBuffer(sizeOfVertex * 5, 0, D3DFVF_XYZRHW Or D3DFVF_DIFFUSE, D3DPOOL_DEFAULT)
    If vb Is Nothing Then Exit Sub
    
    pos.x = (m_currentRoomCol - m_worldCol) * k_ROOM_SPACING_X_PIX
    pos.y = (m_currentRoomRow - m_worldRow) * k_ROOM_SPACING_Y_PIX
    
    If flashOn Then
        r = 255
    Else
        r = 0
    End If
    
    flashOn = Not flashOn
    
    g = 0
    b = 0
    lineList(0).x = pos.x
    lineList(0).y = pos.y
    lineList(0).z = 0
    lineList(0).rhw = 1
    lineList(0).color = D3DColorXRGB(r, g, b)
    
    lineList(1).x = pos.x + k_ROOM_SPACING_X_PIX - 1
    lineList(1).y = pos.y
    lineList(1).z = 0
    lineList(1).rhw = 1
    lineList(1).color = D3DColorXRGB(r, g, b)
    
    lineList(2).x = pos.x + k_ROOM_SPACING_X_PIX - 1
    lineList(2).y = pos.y + k_ROOM_SPACING_Y_PIX - 1
    lineList(2).z = 0
    lineList(2).rhw = 1
    lineList(2).color = D3DColorXRGB(r, g, b)
    
    lineList(3).x = pos.x
    lineList(3).y = pos.y + k_ROOM_SPACING_Y_PIX - 1
    lineList(3).z = 0
    lineList(3).rhw = 1
    lineList(3).color = D3DColorXRGB(r, g, b)
    
    lineList(4).x = pos.x
    lineList(4).y = pos.y
    lineList(4).z = 0
    lineList(4).rhw = 1
    lineList(4).color = D3DColorXRGB(r, g, b)

    D3DVertexBuffer8SetData vb, 0, sizeOfVertex * 5, 0, lineList(0)
    g_d3dDevice.SetStreamSource 0, vb, sizeOfVertex
    g_d3dDevice.SetVertexShader D3DFVF_XYZRHW Or D3DFVF_DIFFUSE
    
    Call g_d3dDevice.DrawPrimitive(D3DPT_LINESTRIP, 0, 4)
End Sub


Private Sub picView_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim col As Integer
    Dim row As Integer

    If Button <> vbLeftButton Then Exit Sub
    
    col = x \ (k_ROOM_SPACING_X_PIX * Screen.TwipsPerPixelX) '\ m_viewHeight
    row = y \ (k_ROOM_SPACING_Y_PIX * Screen.TwipsPerPixelY) 'Index Mod m_viewHeight
    m_currentRoomCol = m_worldCol + col
    m_currentRoomRow = m_worldRow + row
    Call UpdateCurrentRoom
End Sub

Private Sub RoomSelectionLeft()
    m_currentRoomCol = IntMax(m_currentRoomCol - 1, 0)
    Call UpdateCurrentRoom
End Sub

Private Sub RoomSelectionRight()
    m_currentRoomCol = IntMin(m_currentRoomCol + 1, g_map.Width - 1)
    Call UpdateCurrentRoom
End Sub

Private Sub RoomSelectionUp()
    m_currentRoomRow = IntMax(m_currentRoomRow - 1, 0)
    Call UpdateCurrentRoom
End Sub

Private Sub RoomSelectionDown()
    m_currentRoomRow = IntMin(m_currentRoomRow + 1, g_map.Height - 1)
    Call UpdateCurrentRoom
End Sub

Private Sub UpdateCurrentRoom()
    Set m_currentRoom = g_map.GetRoom(m_currentRoomCol, m_currentRoomRow)
    UpdateMapInfo
    UpdateRoomProperties m_currentRoom
End Sub

Private Sub timerAnimation_Timer()
    If m_animateMap Then
        g_frameCount = g_frameCount + 1
    End If
    
    Call RenderScene
End Sub

Private Sub RenderScene()
    Dim d3ddev As Direct3DDevice8
    Dim hr As Long
    Dim room As CRoom
    Dim col As Integer
    Dim row As Integer
    Dim startTick As Long
    Dim endTick As Long
    Dim pos As D3DVECTOR2
    Dim r As RECT
    Dim oldVP As D3DVIEWPORT8
    Dim newVP As D3DVIEWPORT8

    On Local Error Resume Next

    Set d3ddev = g_d3dObject.GetDirect3DDevice
    hr = d3ddev.TestCooperativeLevel
    If hr = D3DERR_DEVICELOST Then
        'If the device is lost, exit and wait for it to come back.
        Exit Sub
    ElseIf hr = D3DERR_DEVICENOTRESET Then
        'The device became lost for some reason (probably an alt-tab) and now
        'Reset() needs to be called to try and get the device back.
        hr = 0
'        hr = ResetDevice()
        'If the device failed to be reset, exit the sub.
        If hr Then Exit Sub
    End If

    'Make sure the app isn't minimized
    If Me.WindowState <> vbMinimized Then
        startTick = GetTickCount
        Call g_d3dDevice.Clear(0, ByVal 0&, D3DCLEAR_TARGET, &H0, 0, 0)
        Call g_d3dDevice.GetViewport(oldVP)
        Call g_d3dDevice.BeginScene
        For col = 0 To m_viewWidth - 1
            For row = 0 To m_viewHeight - 1
                Set room = g_map.GetRoom(m_worldCol + col, m_worldRow + row)
                pos.x = col * k_ROOM_SPACING_X_PIX + 1
                pos.y = row * k_ROOM_SPACING_Y_PIX + 1
                newVP.x = pos.x
                newVP.y = pos.y
                newVP.Width = k_ROOM_WIDTH_PIX
                newVP.Height = k_ROOM_HEIGHT_PIX
                Call g_d3dDevice.SetViewport(newVP)
                Call room.Draw(pos)
            Next

        Next
    
        Call g_d3dDevice.SetViewport(oldVP)
        Call HighlightRoom
        Call g_d3dDevice.EndScene
        endTick = GetTickCount
        UpdateRedrawTime statusBar, "Redraw", endTick - startTick
        r.Left = 0
        r.Top = 0
        r.Right = m_viewWidth * k_ROOM_SPACING_X_PIX
        r.bottom = m_viewHeight * k_ROOM_SPACING_Y_PIX
        
        Call g_d3dDevice.Present(r, r, 0, ByVal 0&)
    End If

End Sub

Private Sub vscrollMap_Change()
    m_worldRow = vscrollMap.value
End Sub

Private Function GetMapMemory() As Long
    GetMapMemory = g_map.Width * g_map.Height * k_ROOM_DATA_SIZE
End Function

Private Sub UpdateMapInfo()
    UpdateRoomNumber
    UpdateStartRoom
    UpdateMemoryStats
End Sub

Private Sub UpdateMemoryStats()
    Dim s As String

    s = "  Memory: " & Trim(GetMapMemory / 1024) & "KB"
    statusBar.Panels.Item("MemoryInfo").Text = s
End Sub

Private Function InputSourceFilenameToExport(ByRef Filename As String) As Boolean
    Filename = InputFilename(".S", "Pitfall 2 Export (*.S)|*.s|All (*.*)|*.*", "Export Pitfall 2 Level As", vbffSaveDialog)
    If Filename = "" Then
        InputSourceFilenameToExport = False
    Else
        InputSourceFilenameToExport = True
    End If

End Function

Private Function InputBinaryFilenameToExport(ByRef Filename As String) As Boolean
    Filename = InputFilename(".BIN", "Pitfall 2 Export (*.BIN)|*.bin|All (*.*)|*.*", "Export Pitfall 2 Level As", vbffSaveDialog)
    If Filename = "" Then
        InputBinaryFilenameToExport = False
    Else
        InputBinaryFilenameToExport = True
    End If

End Function

Private Function InputProjectFilenameToSave(ByRef Filename As String) As Boolean
    Filename = InputFilename(".PF2", "Pitfall 2 Level (*.PF2)|*.pf2|All (*.*)|*.*", "Save Pitfall 2 Level As", vbffSaveDialog)
    If Filename = "" Then
        InputProjectFilenameToSave = False
    Else
        InputProjectFilenameToSave = True
    End If

End Function

Private Function InputProjectFilenameToOpen(ByRef selectedFilename As String) As Boolean
    selectedFilename = InputFilename(".PF2", "Pitfall 2 Level (*.PF2)|*.pf2|All (*.*)|*.*", "Open Pitfall 2 Level File", vbffLoadDialog)
    If selectedFilename = "" Then
        InputProjectFilenameToOpen = False
    Else
        InputProjectFilenameToOpen = True
    End If

End Function

Private Sub CreateNewLevel()
    g_projectFilename = ""
    ' create a new level map
    Set g_map = Nothing
    Set g_map = New CMap
    ' enable save project
    EnableProjectCommands
    InitMapPosition
    UpdateMapInfo
    SetWindowTitle
End Sub

Private Sub LoadLevelFromMRU(ByVal index As Integer)
    Dim ret As Boolean
    Dim Filename As String

    ' get file from mru list
    g_projectFilename = m_mruList.Item(index)
    ' load level from supplied file
    If LoadLevelFromFile(g_projectFilename) Then
        InitMapPosition
        UpdateMapInfo
        SetWindowTitle
        m_mruList.AddItem g_projectFilename
    End If

End Sub

Private Sub LoadLevel()
    Dim ret As Boolean
    Dim Filename As String

    ' get file from user
    ret = InputProjectFilenameToOpen(Filename)
    If Not ret Then
        Exit Sub
    End If

    g_projectFilename = Filename
    ' load level from supplied file
    If LoadLevelFromFile(g_projectFilename) Then
        InitMapPosition
        UpdateMapInfo
        SetWindowTitle
        m_mruList.AddItem g_projectFilename
    End If

End Sub

Private Sub ExportLevel()
'    Dim ret As Boolean
'    Dim Filename As String
'
'    ' use project properties to export the level (if the properties are already set)
'    ' add properties for attaching a user defined label
'    If m_editMode = vbemGameboy Then
'        ret = InputSourceFilenameToExport(Filename)
'    ElseIf m_editMode = vbemVCS Then
'        ret = InputBinaryFilenameToExport(Filename)
'    Else
'        Debug.Assert False
'    End If
'
'    If Not ret Then
'        Exit Sub
'    End If
'
'    If m_editMode = vbemGameboy Then
'        ExportLevelToGameboy Filename
'    ElseIf m_editMode = vbemVCS Then
'        ExportLevelToVCS Filename
'    Else
'        Debug.Assert False
'    End If
End Sub

Private Sub SaveLevel()
    Dim ret As Boolean
    Dim Filename As String

    If g_projectFilename = "" Then
        ret = InputProjectFilenameToSave(Filename)
        If Not ret Then
            Exit Sub
        End If

        g_projectFilename = Filename
    End If

    SaveLevelToFile g_projectFilename
    SetWindowTitle
    m_mruList.AddItem g_projectFilename
End Sub

Private Sub SaveLevelAs()
    Dim ret As Boolean
    Dim Filename As String

    ret = InputProjectFilenameToSave(Filename)
    If Not ret Then
        Exit Sub
    End If

    g_projectFilename = Filename
    SaveLevel
End Sub

Private Sub ExportLevelToVCS(ByVal baseFilename As String)
    ExportBinaryFile baseFilename
End Sub

Private Sub ExportLevelToGameboy(ByVal baseFilename As String)
    Dim labelName As String

    labelName = "LEVEL_1"
    ExportSourceFile baseFilename, labelName
    ExportIncludeFile baseFilename, labelName
End Sub

Private Sub ExportBinaryFile(ByVal baseFilename As String)
'    Dim F As Integer
'    Dim label As String
'    Dim headerFilename As String
'    Dim exportData As CVCSImage
'
'    If g_map.Width <> 8 Then
'        MsgBox "The map must be 8 rooms wide"
'        Exit Sub
'    End If
'
'    If g_map.Height <> 32 Then
'        MsgBox "The map must be 32 rooms high"
'        Exit Sub
'    End If
'
'    On Local Error GoTo ErrorHandler
'    ' create atari binary
'    Set exportData = New CVCSImage
'    g_map.ExportBinary exportData
'
'    ' write source file
'    F = FreeFile
'    Open baseFilename For Binary Access Write As F
'    Put #F, , exportData.GetData
'    Close F
'    Exit Sub
'ErrorHandler:
    MsgBox "Failed to export source file"
End Sub

Private Sub ExportSourceFile(ByVal baseFilename As String, ByVal labelName As String)
    Dim F As Integer
    Dim label As String
    Dim headerFilename As String
    Dim exportData As CExport

    On Local Error GoTo ErrorHandler
    Set exportData = New CExport
    ' create source code
    exportData.Header baseFilename, k_TITLE & " Level Data", g_projectFilename
    g_map.ExportSource exportData, labelName

    ' write source file
    F = FreeFile
    Open baseFilename For Output Access Write As F
    Print #F, exportData.GetData
    Close F
    Exit Sub
ErrorHandler:
    MsgBox "Failed to export source file"
End Sub

Private Sub ExportIncludeFile(ByVal baseFilename As String, ByVal labelName As String)
    Dim F As Integer
    Dim label As String
    Dim headerFilename As String
    Dim exportData As CExport

    On Local Error GoTo ErrorHandler
    Set exportData = New CExport
    ' remove .s and replace with .i
    headerFilename = RemoveExtension(baseFilename) & ".i"

    ' create header code
    exportData.Reset
    exportData.Header headerFilename, k_TITLE & " Level Data", g_projectFilename
    exportData.ConditionalInclude labelName
    g_map.ExportHeader exportData, labelName
    exportData.ConditionalEnd

    ' write header file
    F = FreeFile
    Open headerFilename For Output Access Write As F
    Print #F, exportData.GetData
    Close F
    Exit Sub
ErrorHandler:
    MsgBox "Failed to export include file"
End Sub

Private Sub Action_PlayGame()
    frmGame.Show
End Sub

Private Sub Action_CutRoom()
    Action_CopyRoom
    Action_ResetRoom
End Sub

Private Sub Action_CopyRoom()
    Set m_copyRoom = New CRoom
    m_copyRoom.Copy m_currentRoom
    EnablePaste
End Sub

Private Sub Action_PasteRoom()
    m_currentRoom.Copy m_copyRoom
    UpdateRoomProperties m_currentRoom
End Sub

Private Sub DisablePaste()
    menuPaste.Enabled = False
    toolbarMain.Buttons.Item("Paste").Enabled = False
End Sub

Private Sub EnablePaste()
    menuPaste.Enabled = True
    toolbarMain.Buttons.Item("Paste").Enabled = True
End Sub

Private Sub Action_ResetRoom()
    m_currentRoom.Reset
    UpdateRoomProperties m_currentRoom
End Sub

Private Sub Action_AnimateMap()
    Debug.Assert m_animateMap = vbmaPause Or m_animateMap = vbmaPlay
    If m_animateMap = vbmaPlay Then
        m_animateMap = vbmaPause
        toolbarMain.Buttons.Item("Animate").Image = "Animate"
    Else
        m_animateMap = vbmaPlay
        toolbarMain.Buttons.Item("Animate").Image = "Pause"
    End If

End Sub

Private Sub vscrollMap_Scroll()
    m_worldRow = vscrollMap.value
End Sub

Private Sub InitToolbars()
    InitToolbarMain
End Sub

Private Sub InitToolbarMain()
    Dim buttonItem As Button

    toolbarMain.BorderStyle = ccNone
    toolbarMain.ImageList = imagelistToolbar
    For Each buttonItem In toolbarMain.Buttons
        If buttonItem.Style <> tbrSeparator Then
            buttonItem.Image = buttonItem.Key
            buttonItem.Caption = ""
        End If
    Next

    toolbarMain.Buttons("Animate").Image = "Pause"
    DisableToolbar
End Sub

Private Sub EnableToolbar()
    Dim buttonItem As Button

    ' enable all buttons
    For Each buttonItem In toolbarMain.Buttons
        buttonItem.Enabled = True
    Next

End Sub

Private Sub DisableToolbar()
    Dim buttonItem As Button

    ' disable all buttons
    For Each buttonItem In toolbarMain.Buttons
        buttonItem.Enabled = False
    Next

    ' enable those buttons that are always enabled
    toolbarMain.Buttons.Item("New").Enabled = True
    toolbarMain.Buttons.Item("Open").Enabled = True
    toolbarMain.Buttons.Item("Copy").Enabled = True
    toolbarMain.Buttons.Item("Cut").Enabled = True
    toolbarMain.Buttons.Item("Animate").Enabled = True
    toolbarMain.Buttons.Item("Play").Enabled = True
'    toolbarMain.Buttons.Item("Undo").Enabled = True
'    toolbarMain.Buttons.Item("Redo").Enabled = True
'    toolbarMain.Buttons.Item("Toolbox").Enabled = True
'    toolbarMain.Buttons.Item("Spawnpoints").Enabled = True
End Sub

Private Sub toolbarMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Key = "New" Then
        VerifiedNewLevel
    ElseIf Button.Key = "Save" Then
        SaveLevel
    ElseIf Button.Key = "Open" Then
        VerifiedLoadLevel
    ElseIf Button.Key = "Cut" Then
        Action_CutRoom
    ElseIf Button.Key = "Copy" Then
        Action_CopyRoom
    ElseIf Button.Key = "Paste" Then
        Action_PasteRoom
    ElseIf Button.Key = "Undo" Then
        Debug.Assert False
    ElseIf Button.Key = "Redo" Then
        Debug.Assert False
    ElseIf Button.Key = "Animate" Then
        Action_AnimateMap
    ElseIf Button.Key = "Play" Then
        Action_PlayGame
    End If

End Sub

Private Sub VerifiedLoadLevel()
    ' if current project has changed
    If g_map.HasChanged Then
            ' if user didn't really want to exit
            If Not VerifyOpenWithoutSaving Then
                ' cancel exit
                ' exit function
                Exit Sub
            End If
    End If

    DisablePaste
    LoadLevel
End Sub

Private Sub VerifiedLoadLevelFromMRU(ByVal index As Integer)
    ' if current project has changed
    If g_map.HasChanged Then
            ' if user didn't really want to exit
            If Not VerifyOpenWithoutSaving Then
                ' cancel exit
                ' exit function
                Exit Sub
            End If
    End If

    DisablePaste
    LoadLevelFromMRU index
End Sub

'Private Sub InitMenuIcons()
'    Dim dcMemory As Long
'    Dim hMemoryBitmap As Long
'    Dim dummy As Long
'
'    Picture1.ScaleMode = vbPixels
'    Picture1.Picture = imagelistToolbar.ListImages.Item(1).Picture
'    Picture1.Refresh
'    dcMemory = CreateCompatibleDC(Picture1.hDC)
'    hMemoryBitmap = CreateCompatibleBitmap(Picture1.hDC, Picture1.ScaleWidth, Picture1.ScaleHeight)
'
'    Dim pObject As Long
'    pObject = SelectObject(dcMemory, hMemoryBitmap)
'    dummy = BitBlt(dcMemory, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.hDC, 0, 0, &HCC0020)
'    dummy = SelectObject(dcMemory, pObject)
'    dummy = ModifyMenu(GetSubMenu(GetMenu(Me.hwnd), 0), 0, MF_BYPOSITION Or MF_USECHECKBITMAPS, 0, hMemoryBitmap)
'End Sub


