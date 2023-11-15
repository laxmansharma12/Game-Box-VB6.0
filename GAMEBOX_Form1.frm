VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000003&
   Caption         =   "Form1"
   ClientHeight    =   10485
   ClientLeft      =   2820
   ClientTop       =   1155
   ClientWidth     =   15870
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10485
   ScaleWidth      =   15870
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   9015
      Left            =   5640
      TabIndex        =   7
      Top             =   1200
      Width           =   4575
      Begin VB.Label Label4 
         Caption         =   "MAZE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1560
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Notes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   8
         Top             =   1080
         Width           =   735
      End
      Begin VB.Line Line4 
         X1              =   2640
         X2              =   4560
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line3 
         X1              =   -120
         X2              =   1680
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Image Image8 
         Height          =   7815
         Left            =   0
         Picture         =   "GAMEBOX_Form1.frx":0000
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   4815
      End
   End
   Begin VB.Frame Frame1 
      Height          =   9015
      Left            =   5640
      TabIndex        =   4
      Top             =   1200
      Width           =   4575
      Begin VB.Image Image7 
         Height          =   7815
         Left            =   0
         Picture         =   "GAMEBOX_Form1.frx":FCEA
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   4455
      End
      Begin VB.Line Line1 
         X1              =   -120
         X2              =   1680
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line2 
         X1              =   2640
         X2              =   4560
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label2 
         Caption         =   "Notes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   6
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Black Jack"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   960
         TabIndex        =   5
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "TIC-TAC-TOE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12480
      TabIndex        =   3
      Top             =   6840
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BLACK-JACK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      TabIndex        =   2
      Top             =   3480
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "MAZE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12480
      TabIndex        =   1
      Top             =   3480
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "QUIZ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      TabIndex        =   0
      Top             =   6840
      Width           =   2415
   End
   Begin VB.Image Image6 
      Height          =   855
      Left            =   6240
      Picture         =   "GAMEBOX_Form1.frx":23F59
      Stretch         =   -1  'True
      Top             =   360
      Width           =   3615
   End
   Begin VB.Image Image5 
      Height          =   1935
      Left            =   12480
      Picture         =   "GAMEBOX_Form1.frx":2BCF9
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   2295
   End
   Begin VB.Image Image4 
      Height          =   1965
      Left            =   1200
      Picture         =   "GAMEBOX_Form1.frx":67969
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   2235
   End
   Begin VB.Image Image2 
      Height          =   1935
      Left            =   12480
      Picture         =   "GAMEBOX_Form1.frx":9B502
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Image Image3 
      Height          =   1935
      Left            =   1200
      Picture         =   "GAMEBOX_Form1.frx":B79B7
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   10440
      Left            =   0
      Picture         =   "GAMEBOX_Form1.frx":D6484
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15840
   End
   Begin VB.Menu fall 
      Caption         =   "&All"
      Begin VB.Menu aedit 
         Caption         =   "&Edit"
         Index           =   0
         Begin VB.Menu aaddplayer 
            Caption         =   "&Add Player"
         End
         Begin VB.Menu aupdateplayer 
            Caption         =   "&Update Player"
         End
         Begin VB.Menu adeleteplayer 
            Caption         =   "&Delete Player"
         End
         Begin VB.Menu asearchplayer 
            Caption         =   "&Search Player"
         End
      End
      Begin VB.Menu ilogout 
         Caption         =   "&Log Out"
      End
      Begin VB.Menu iexit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu fedit 
      Caption         =   "&Edit"
      Begin VB.Menu eaddplayer 
         Caption         =   "&Add Player"
      End
      Begin VB.Menu eupdateplayer 
         Caption         =   "&Update Player"
      End
      Begin VB.Menu edeleteplayer 
         Caption         =   "&Delete Player"
      End
      Begin VB.Menu esearchplayer 
         Caption         =   "&Search Player"
      End
   End
   Begin VB.Menu fview 
      Caption         =   "&View"
      Begin VB.Menu iallplayer 
         Caption         =   "&All Player"
      End
      Begin VB.Menu iscores 
         Caption         =   "&Scores"
         Begin VB.Menu igblackjack 
            Caption         =   "&Black Jack"
         End
         Begin VB.Menu igmaze 
            Caption         =   "&Maze"
         End
         Begin VB.Menu igquiz 
            Caption         =   "&Quiz"
         End
         Begin VB.Menu igtictactoe 
            Caption         =   "&Tic Tac Toe"
         End
      End
   End
   Begin VB.Menu frun 
      Caption         =   "&Play"
      Begin VB.Menu pblackjack 
         Caption         =   "&Black Jack"
      End
      Begin VB.Menu pmaze 
         Caption         =   "&Maze"
      End
      Begin VB.Menu pquiz 
         Caption         =   "&Quiz"
      End
      Begin VB.Menu ptictactoe 
         Caption         =   "&Tic Tac Toe"
      End
   End
   Begin VB.Menu fhelp 
      Caption         =   "&Help"
      Begin VB.Menu gblackjack 
         Caption         =   "&Black Jack"
      End
      Begin VB.Menu gmaze 
         Caption         =   "&Maze"
      End
      Begin VB.Menu gquiz 
         Caption         =   "&Quiz"
      End
      Begin VB.Menu gtictactoe 
         Caption         =   "&Tic Tac Toe"
      End
   End
   Begin VB.Menu fexit 
      Caption         =   "&Exit"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Visible = True
Unload Me
End Sub

Private Sub Command2_Click()
Form10.Visible = True
Unload Me
End Sub

Private Sub Command3_Click()
Form11.Visible = True
Unload Me
End Sub

Private Sub Command5_Click()
Form13.Visible = True
Unload Me
End Sub

Private Sub aaddplayer_Click()
Load Form8
Form8.Show
Unload Me
End Sub

Private Sub adeleteplayer_Click()
Load Form15
Form15.Show
Unload Me
End Sub

Private Sub asearchplayer_Click()
Load Form16
Form16.Show
Unload Me
End Sub


Private Sub aupdateplayer_Click()
Load Form14
Form14.Show
Unload Me
End Sub
Private Sub eaddplayer_Click()
Load Form8
Form8.Show
Unload Me
End Sub

Private Sub edeleteplayer_Click()
Load Form15
Form15.Show
Unload Me
End Sub

Private Sub esearchplayer_Click()
Load Form16
Form16.Show
Unload Me
End Sub


Private Sub eupdateplayer_Click()
Load Form14
Form14.Show
Unload Me
End Sub


Private Sub fexit_Click()
End
End Sub

Private Sub iallplayer_Click()
Load Form17
Form17.Show
Unload Me
End Sub

Private Sub iexit_Click()
End
End Sub

Private Sub ilogout_Click()
Load Form7
Form7.Show
Unload Me
End Sub

Private Sub pblackjack_Click()
Load Form2
Form2.Show
Unload Me
End Sub

Private Sub pmaze_Click()
Load Form10
Form10.Show
Unload Me
End Sub

Private Sub pquiz_Click()
Load Form11
Form11.Show
Unload Me
End Sub

Private Sub ptictactoe_Click()
Load Form13
Form13.Show
Unload Me
End Sub
