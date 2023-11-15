VERSION 5.00
Begin VB.Form Form12 
   Caption         =   "Form12"
   ClientHeight    =   6960
   ClientLeft      =   4365
   ClientTop       =   2580
   ClientWidth     =   13800
   LinkTopic       =   "Form12"
   ScaleHeight     =   6960
   ScaleWidth      =   13800
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "FINISH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   2
      Top             =   6120
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   9480
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      Begin VB.Timer Timer1 
         Interval        =   900
         Left            =   3720
         Top             =   1080
      End
      Begin VB.Label Label48 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   30
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   1560
         TabIndex        =   1
         Top             =   2640
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C00000&
      Height          =   255
      Left            =   0
      TabIndex        =   49
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C00000&
      Height          =   6855
      Left            =   0
      TabIndex        =   48
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C00000&
      Height          =   255
      Left            =   240
      TabIndex        =   47
      Top             =   6720
      Width           =   8895
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C00000&
      Height          =   6615
      Left            =   9120
      TabIndex        =   46
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C00000&
      Height          =   255
      Left            =   6600
      TabIndex        =   45
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C00000&
      Height          =   255
      Left            =   240
      TabIndex        =   44
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C00000&
      Height          =   2175
      Left            =   3240
      TabIndex        =   43
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C00000&
      Height          =   2175
      Left            =   1320
      TabIndex        =   42
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C00000&
      Height          =   255
      Left            =   3240
      TabIndex        =   41
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C00000&
      Height          =   255
      Left            =   3240
      TabIndex        =   40
      Top             =   4320
      Width           =   2775
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C00000&
      Height          =   255
      Left            =   4680
      TabIndex        =   39
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C00000&
      Height          =   1215
      Left            =   6600
      TabIndex        =   38
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C00000&
      Height          =   255
      Left            =   600
      TabIndex        =   37
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C00000&
      Height          =   1575
      Left            =   2280
      TabIndex        =   36
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label15 
      BackColor       =   &H00C00000&
      Height          =   1215
      Left            =   1680
      TabIndex        =   35
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C00000&
      Height          =   255
      Left            =   1080
      TabIndex        =   34
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label Label17 
      BackColor       =   &H00C00000&
      Height          =   255
      Left            =   1440
      TabIndex        =   33
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label18 
      BackColor       =   &H00C00000&
      Height          =   255
      Left            =   3960
      TabIndex        =   32
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label19 
      BackColor       =   &H00C00000&
      Height          =   975
      Left            =   1800
      TabIndex        =   31
      Top             =   5400
      Width           =   255
   End
   Begin VB.Label Label20 
      BackColor       =   &H00C00000&
      Height          =   255
      Left            =   1800
      TabIndex        =   30
      Top             =   6120
      Width           =   2775
   End
   Begin VB.Label Label21 
      BackColor       =   &H00C00000&
      Height          =   975
      Left            =   4320
      TabIndex        =   29
      Top             =   5400
      Width           =   255
   End
   Begin VB.Label Label22 
      BackColor       =   &H00C00000&
      Height          =   255
      Left            =   4320
      TabIndex        =   28
      Top             =   5160
      Width           =   3375
   End
   Begin VB.Label Label23 
      BackColor       =   &H00C00000&
      Height          =   855
      Left            =   5040
      TabIndex        =   27
      Top             =   6000
      Width           =   255
   End
   Begin VB.Label Label24 
      BackColor       =   &H00C00000&
      Height          =   255
      Left            =   5040
      TabIndex        =   26
      Top             =   5760
      Width           =   2535
   End
   Begin VB.Label Label25 
      BackColor       =   &H00C00000&
      Height          =   2655
      Left            =   7440
      TabIndex        =   25
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C00000&
      Height          =   255
      Left            =   5160
      TabIndex        =   24
      Top             =   2400
      Width           =   2535
   End
   Begin VB.Label Label27 
      BackColor       =   &H00C00000&
      Height          =   1095
      Left            =   5640
      TabIndex        =   23
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label28 
      BackColor       =   &H00C00000&
      Height          =   975
      Left            =   5640
      TabIndex        =   22
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label29 
      BackColor       =   &H00C00000&
      Height          =   255
      Left            =   4920
      TabIndex        =   21
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label30 
      BackColor       =   &H00C00000&
      Height          =   975
      Left            =   4320
      TabIndex        =   20
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label31 
      BackColor       =   &H00C00000&
      Height          =   735
      Left            =   2160
      TabIndex        =   19
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label32 
      BackColor       =   &H00C00000&
      Height          =   255
      Left            =   2280
      TabIndex        =   18
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label Label33 
      BackColor       =   &H00C00000&
      Height          =   1575
      Left            =   6600
      TabIndex        =   17
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label34 
      BackColor       =   &H00C00000&
      Height          =   255
      Left            =   6120
      TabIndex        =   16
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label35 
      BackColor       =   &H00C00000&
      Height          =   735
      Left            =   5040
      TabIndex        =   15
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label Label36 
      BackColor       =   &H00C00000&
      Height          =   255
      Left            =   3480
      TabIndex        =   14
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label37 
      BackColor       =   &H00C00000&
      Height          =   975
      Left            =   3240
      TabIndex        =   13
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label38 
      BackColor       =   &H00C00000&
      Height          =   2175
      Left            =   8280
      TabIndex        =   12
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label39 
      BackColor       =   &H00C00000&
      Height          =   255
      Left            =   7320
      TabIndex        =   11
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label40 
      BackColor       =   &H00C00000&
      Height          =   255
      Left            =   7680
      TabIndex        =   10
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label41 
      BackColor       =   &H00C00000&
      Height          =   855
      Left            =   6240
      TabIndex        =   9
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label Label42 
      BackColor       =   &H00C00000&
      Height          =   255
      Left            =   8160
      TabIndex        =   8
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label43 
      BackColor       =   &H00C00000&
      Height          =   255
      Left            =   7440
      TabIndex        =   7
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Label44 
      BackColor       =   &H00C00000&
      Height          =   1215
      Left            =   8760
      TabIndex        =   6
      Top             =   3120
      Width           =   255
   End
   Begin VB.Label Label45 
      BackColor       =   &H00C00000&
      Height          =   975
      Left            =   8280
      TabIndex        =   5
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label Label46 
      BackColor       =   &H00C00000&
      Height          =   255
      Left            =   7920
      TabIndex        =   4
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label Label47 
      BackColor       =   &H00C00000&
      Height          =   255
      Left            =   8160
      TabIndex        =   3
      Top             =   6240
      Width           =   1095
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MsgBox "You Win"
Form10.Visible = True
Unload Me
End Sub

Private Sub Label48_Click()
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
Label48.Caption = Label48.Caption + 1
End Sub
Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label14_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label15_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label16_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label17_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label18_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label19_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label20_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label21_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label22_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label23_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label24_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label25_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label26_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label27_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label28_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label29_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label30_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label31_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label32_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label33_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label34_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label35_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label36_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label37_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label38_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label39_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label40_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label41_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label42_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label43_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label44_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label45_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label46_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label47_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub

