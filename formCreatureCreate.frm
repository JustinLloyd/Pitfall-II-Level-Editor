VERSION 5.00
Begin VB.Form formCreatureEditor 
   Caption         =   "Creature Editor"
   ClientHeight    =   6840
   ClientLeft      =   6840
   ClientTop       =   5580
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   9405
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse..."
      Height          =   375
      Index           =   0
      Left            =   5880
      TabIndex        =   10
      Top             =   3000
      Width           =   1095
   End
   Begin VB.PictureBox picAnimFrame 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   0
      Left            =   7200
      ScaleHeight     =   615
      ScaleWidth      =   855
      TabIndex        =   9
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox textBitmapFilepath 
      Height          =   285
      Index           =   0
      Left            =   1320
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   3000
      Width           =   4455
   End
   Begin VB.PictureBox picRoom 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      ScaleHeight     =   975
      ScaleWidth      =   3255
      TabIndex        =   6
      Top             =   120
      Width           =   3255
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3600
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   360
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   8040
      TabIndex        =   3
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Accept"
      Default         =   -1  'True
      Height          =   375
      Left            =   6480
      TabIndex        =   2
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Create"
      Height          =   495
      Left            =   7320
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Remove"
      Height          =   495
      Left            =   7320
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label labelFrameNum 
      Caption         =   "Label2"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   8
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Flight Path"
      Height          =   255
      Left            =   3600
      TabIndex        =   5
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "formCreatureEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' TO DO
'=======
' Add Y offset for creature
' Add flightpath for creature
' Add "does it home" for creature
' Add start pos (X & Y) for creature
' Add load bitmaps
' Add create new creature
' Add remove creature

' Flight paths are:
' start at left, start at right
' haphazard path 1
' slow sinewave path
' fast sinewave path


Private m_room As CRoom

Private Sub Form_Load()
    Initialise
End Sub

Private Sub InitRoom()
'    Set m_room = New CRoom
'
''    m_room.Background = k_BACKGROUND_NONE
'    m_room.Floor = k_FLOOR_WALKWAY
'    picRoom.ScaleMode = vbTwips
''    x = k_ROOM_WIDTH_PIX + 2
''    x = startX + x * Screen.TwipsPerPixelX
''    y = k_ROOM_HEIGHT_PIX + 2
''    y = startY + y * Screen.TwipsPerPixelY
'    picRoom.Visible = True
''    room.Left = x
''    room.Top = y
'    picRoom.Height = Screen.TwipsPerPixelY * k_ROOM_HEIGHT_PIX
'    picRoom.Width = Screen.TwipsPerPixelX * k_ROOM_WIDTH_PIX
'    picRoom.ScaleMode = vbPixels
''    DrawRoom picRoom, m_room
End Sub

Private Sub Initialise()
    Dim index As Integer
    
    For index = 0 To k_MAX_ANIM_FRAMES - 1
        If index <> 0 Then
            Load textBitmapFilepath(index)
            Load labelFrameNum(index)
            Load cmdBrowse(index)
            Load picAnimFrame(index)
        End If
        
        textBitmapFilepath(index).Text = ""
        labelFrameNum(index).Caption = "Frame #1"
        textBitmapFilepath(index).Top = textBitmapFilepath(0).Top + (textBitmapFilepath(0).Height + 100) * index
        textBitmapFilepath(index).Visible = True
        labelFrameNum(index).Visible = True
    Next
    
    InitRoom
End Sub
