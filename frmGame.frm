VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGame 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Game"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5130
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   5130
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picRoom 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      ScaleHeight     =   615
      ScaleWidth      =   2055
      TabIndex        =   2
      Top             =   2520
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.PictureBox picBackbuffer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2025
      Left            =   2640
      ScaleHeight     =   2025
      ScaleWidth      =   2115
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   120
      ScaleHeight     =   2175
      ScaleWidth      =   2055
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin MSComctlLib.StatusBar statusBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   3
      Top             =   3945
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
            Key             =   "MapInfo"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "Redraw"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const k_SCREEN_WIDTH = 160&
Private Const k_SCREEN_HEIGHT = 144&
Private Const k_SCREEN_SCALE = 2&

Private Const k_MSG_GAME_PAUSED = "Game Paused"
Private Const k_MSG_GAME_RUNNING = "Game In Progress"

Private Const vbPlayerKeyLeft = vbKeyLeft
Private Const vbPlayerKeyRight = vbKeyRight
Private Const vbPlayerKeyUp = vbKeyUp
Private Const vbPlayerKeyDown = vbKeyDown
Private Const vbPlayerKeyJump = vbKeyA

Private Const k_KEY_UP = 1
Private Const k_KEY_DOWN = 2
Private Const k_KEY_LEFT = 4
Private Const k_KEY_RIGHT = 8
Private Const k_KEY_JUMP = 16

Private Const k_SCREEN_CENTRE_Y = k_SCREEN_HEIGHT \ 2
Private Const k_SCROLL_BOUNDARY = 40

Private m_drawTime As Long
Private m_isPlaying As Boolean
Private m_frameCount As Long
Private m_keyPress As Long

Private m_mapX As Long
Private m_mapY As Long

Private Type HarryType
    m_screenX As Long
    m_screenY As Long
    m_worldX As Long
    m_worldY As Long
    m_deltaY As Long
    m_deltaX As Long
    m_worldPreviousX As Long
    m_worldPreviousY As Long
    m_mapCol As Long
    m_mapRow As Long
End Type

Private m_harry As HarryType



Private Sub Form_Load()
    Initialise
End Sub

Private Sub InitRoomBuffer()
    picRoom.ScaleMode = vbTwips
'    picRoom.Visible = False
    picRoom.Height = Screen.TwipsPerPixelY * k_ROOM_HEIGHT_PIX
    picRoom.Width = Screen.TwipsPerPixelX * k_ROOM_WIDTH_PIX
    picRoom.ScaleMode = vbPixels
'    picRoom.BorderStyle = vbBSNone
End Sub

Private Sub InitPlayer()
'    Dim x As Long
'    Dim y As Long
'    Dim col As Long
'    Dim row As Long
'
'    m_harry.m_deltaX = 1
'    m_harry.m_deltaY = 0
'    m_harry.m_mapCol = g_map.StartRoom \ g_map.Height
'    m_harry.m_mapRow = g_map.StartRoom Mod g_map.Height
'    m_harry.m_worldX = m_harry.m_mapCol * k_ROOM_WIDTH_PIX + k_FEATURE_SAVE_POINT_X
'    m_harry.m_worldY = m_harry.m_mapRow * k_ROOM_HEIGHT_PIX + k_FEATURE_SAVE_POINT_Y
''    m_harry.m_screenX = k_FEATURE_SAVE_POINT_X
''    m_harry.m_screenY = k_FEATURE_SAVE_POINT_Y
'    m_harry.m_worldPreviousX = m_harry.m_worldX
'    m_harry.m_worldPreviousY = m_harry.m_worldY
'    CalculatePlayerScreenPos
End Sub

Private Sub Initialise()
    InitPlayer
    InitRoomBuffer
    m_mapY = 0
    m_mapX = 0
    ' set the on-screen buffer to the correct size
    picScreen.Move 0, 0, LCDScreenWidth(), LCDScreenHeight()
    ' set up the back buffer
'    picBackbuffer.Visible = False
    picBackbuffer.Width = LCDScreenWidth()
    ' make the back buffer tall enough to have 5 rooms drawn to it
    picBackbuffer.Height = Screen.TwipsPerPixelY * k_ROOM_HEIGHT_PIX * 5
    
    m_isPlaying = False
    m_frameCount = 0
    UpdateRedrawTime statusBar, "Redraw", 0
'    InsizeForm Me, picScreen.Width, picScreen.Height + statusBar.Height
    CenterForm Me
End Sub

Private Function LCDScreenWidth()
    LCDScreenWidth = k_SCREEN_WIDTH * Screen.TwipsPerPixelX
End Function

Private Function LCDScreenHeight()
    LCDScreenHeight = k_SCREEN_HEIGHT * Screen.TwipsPerPixelY
End Function

Private Sub GetPlayerInput()
    If GetAsyncKeyState(vbPlayerKeyLeft) Then
        m_keyPress = m_keyPress Or k_KEY_LEFT
    End If
    If GetAsyncKeyState(vbPlayerKeyRight) Then
        m_keyPress = m_keyPress Or k_KEY_RIGHT
    End If
    If GetAsyncKeyState(vbPlayerKeyUp) Then
        m_keyPress = m_keyPress Or k_KEY_UP
    End If
    If GetAsyncKeyState(vbPlayerKeyDown) Then
        m_keyPress = m_keyPress Or k_KEY_DOWN
    End If
    If GetAsyncKeyState(vbPlayerKeyJump) Then
        m_keyPress = m_keyPress Or k_KEY_JUMP
    End If
End Sub

Private Sub DrawScreen()
    Dim ret As Long
    
'    DrawPlayer
'    DrawScore
'    DrawDeveloperLogo
    ret = BitBlt(picScreen.hdc, 0, 0, k_SCREEN_WIDTH, k_SCREEN_HEIGHT, picBackbuffer.hdc, 0&, m_mapY Mod k_ROOM_HEIGHT_PIX, SRCCOPY)
End Sub

Private Sub DrawPlayer()
'    DrawPlayerSprite picBackbuffer.hdc, m_harry.m_screenX, m_harry.m_screenY + m_mapY Mod k_ROOM_HEIGHT_PIX
End Sub

Private Sub PlayGame()
    Do
        m_frameCount = m_frameCount + 1
        DrawScreen
        GetPlayerInput
        UpdatePlayer
        ScrollMap
        ' this is here to stop the animation getting too fast to see
        Sleep 33 - m_drawTime
        ' ensure we can still click buttons etc
        DoEvents
    Loop While m_isPlaying
End Sub

Private Sub DrawScore()

End Sub

Private Sub DrawDeveloperLogo()

End Sub

Private Sub ScrollMap()
    Dim diff As Long
    
    diff = m_harry.m_worldY - m_mapY
    If diff < k_SCREEN_CENTRE_Y - k_SCROLL_BOUNDARY Then
        m_mapY = m_mapY - m_harry.m_deltaY
    ElseIf diff > k_SCREEN_CENTRE_Y + k_SCROLL_BOUNDARY Then
        m_mapY = m_mapY + m_harry.m_deltaY
    End If
    
    ' is player in an exit area?
    diff = m_harry.m_worldX - m_mapX
    If diff < 10 Then
        ' is the player moving left?
        If m_harry.m_deltaX < 0 Then
            m_mapX = m_mapX - k_SCREEN_WIDTH
        End If
    ElseIf diff > k_SCREEN_WIDTH - 10 Then
        ' is the player moving right?
        If m_harry.m_deltaX > 0 Then
            m_mapX = m_mapX + k_SCREEN_WIDTH
        End If
    End If
    
    ' has the player attempted to move off the map?
    If m_mapX < 0 Then m_mapX = 0
    If m_mapX > g_map.Width * k_SCREEN_WIDTH - 16 Then m_mapX = (g_map.Width - 1) * k_SCREEN_HEIGHT - 16
    If m_mapY < 0 Then m_mapY = 0
    If m_mapY > g_map.Height * k_SCREEN_HEIGHT - 16 Then m_mapY = (g_map.Height - 3) * k_SCREEN_HEIGHT - 16
    

End Sub

Private Sub UpdatePlayer()
    m_harry.m_worldPreviousX = m_harry.m_worldX
    m_harry.m_worldPreviousY = m_harry.m_worldY
    m_harry.m_worldX = m_harry.m_worldX + m_harry.m_deltaX
    m_harry.m_worldY = m_harry.m_worldY + m_harry.m_deltaY
    
    If m_harry.m_worldX < 0 Then m_harry.m_worldX = 0
    If m_harry.m_worldX > g_map.Width * k_SCREEN_WIDTH Then m_harry.m_worldX = g_map.Width * k_SCREEN_WIDTH
    If m_harry.m_worldY < 0 Then m_harry.m_worldY = 0
    If m_harry.m_worldY > g_map.Height * k_SCREEN_HEIGHT Then m_harry.m_worldY = g_map.Height * k_SCREEN_HEIGHT

    CalculatePlayerScreenPos
End Sub

Private Sub CalculatePlayerScreenPos()
    m_harry.m_screenY = m_harry.m_worldY - m_mapY
    m_harry.m_screenX = m_harry.m_worldX - m_mapX
End Sub

Private Sub SetWindowTitle(ByVal title As String)
    Me.Caption = title
End Sub

Private Sub Action_TogglePlay()
    If m_isPlaying Then
        m_isPlaying = False
        SetWindowTitle k_MSG_GAME_PAUSED
    Else
        m_isPlaying = True
        SetWindowTitle k_MSG_GAME_RUNNING
        PlayGame
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    m_isPlaying = False
    DoEvents
    Sleep 50
    DoEvents
End Sub

Private Sub picScreen_Click()
    Action_TogglePlay
End Sub

Public Sub PauseGame()
    m_isPlaying = False
    SetWindowTitle k_MSG_GAME_PAUSED
End Sub
