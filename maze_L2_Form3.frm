VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H8000000E&
   Caption         =   "Form3"
   ClientHeight    =   8295
   ClientLeft      =   4755
   ClientTop       =   4320
   ClientWidth     =   16530
   LinkTopic       =   "Form3"
   ScaleHeight     =   8295
   ScaleWidth      =   16530
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
      Height          =   375
      Left            =   5760
      TabIndex        =   2
      Top             =   7680
      Width           =   1335
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
      Height          =   8055
      Left            =   12120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.Timer Timer1 
         Interval        =   900
         Left            =   3720
         Top             =   1080
      End
      Begin VB.Label Label41 
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
         Height          =   1215
         Left            =   1680
         TabIndex        =   1
         Top             =   2880
         Width           =   1095
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C000&
      Height          =   375
      Left            =   1800
      TabIndex        =   51
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   375
      Left            =   120
      TabIndex        =   50
      Top             =   6120
      Width           =   3255
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   375
      Left            =   1800
      TabIndex        =   49
      Top             =   7560
      Width           =   2895
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C000&
      Height          =   375
      Left            =   0
      TabIndex        =   48
      Top             =   2640
      Width           =   3255
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C000&
      Height          =   375
      Left            =   2160
      TabIndex        =   47
      Top             =   5280
      Width           =   2415
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C000&
      Height          =   3975
      Left            =   0
      TabIndex        =   46
      Top             =   1320
      Width           =   135
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C000&
      Height          =   2055
      Left            =   3720
      TabIndex        =   45
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000C000&
      Height          =   375
      Left            =   0
      TabIndex        =   44
      Top             =   960
      Width           =   3255
   End
   Begin VB.Label Label9 
      BackColor       =   &H0000C000&
      Height          =   375
      Left            =   1800
      TabIndex        =   43
      Top             =   6960
      Width           =   2415
   End
   Begin VB.Label Label10 
      BackColor       =   &H0000C000&
      Height          =   2295
      Left            =   3240
      TabIndex        =   42
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label11 
      BackColor       =   &H0000C000&
      Height          =   1935
      Left            =   2160
      TabIndex        =   41
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label12 
      BackColor       =   &H0000C000&
      Height          =   1935
      Left            =   4200
      TabIndex        =   40
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label13 
      BackColor       =   &H0000C000&
      Height          =   3015
      Left            =   0
      TabIndex        =   39
      Top             =   5280
      Width           =   135
   End
   Begin VB.Label Label14 
      BackColor       =   &H0000C000&
      Height          =   1935
      Left            =   3840
      TabIndex        =   38
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label Label15 
      BackColor       =   &H0000C000&
      Height          =   1935
      Left            =   4800
      TabIndex        =   37
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label Label16 
      BackColor       =   &H0000C000&
      Height          =   1935
      Left            =   4440
      TabIndex        =   36
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label17 
      BackColor       =   &H0000C000&
      Height          =   375
      Left            =   4080
      TabIndex        =   35
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label18 
      BackColor       =   &H0000C000&
      Height          =   375
      Left            =   4800
      TabIndex        =   34
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Label Label19 
      BackColor       =   &H0000C000&
      Height          =   375
      Left            =   4440
      TabIndex        =   33
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label20 
      BackColor       =   &H0000C000&
      Height          =   1575
      Left            =   6240
      TabIndex        =   32
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label21 
      BackColor       =   &H0000C000&
      Height          =   2175
      Left            =   7920
      TabIndex        =   31
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label22 
      BackColor       =   &H0000C000&
      Height          =   135
      Left            =   2400
      TabIndex        =   30
      Top             =   0
      Width           =   7575
   End
   Begin VB.Label Label23 
      BackColor       =   &H0000C000&
      Height          =   3015
      Left            =   6840
      TabIndex        =   29
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label24 
      BackColor       =   &H0000C000&
      Height          =   375
      Left            =   6840
      TabIndex        =   28
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label25 
      BackColor       =   &H0000C000&
      Height          =   375
      Left            =   7680
      TabIndex        =   27
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label26 
      BackColor       =   &H0000C000&
      Height          =   1935
      Left            =   7440
      TabIndex        =   26
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label Label27 
      BackColor       =   &H0000C000&
      Height          =   375
      Left            =   6000
      TabIndex        =   25
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label Label28 
      BackColor       =   &H0000C000&
      Height          =   495
      Left            =   9000
      TabIndex        =   24
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label29 
      BackColor       =   &H0000C000&
      Height          =   2295
      Left            =   8280
      TabIndex        =   23
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label Label30 
      BackColor       =   &H0000C000&
      Height          =   2415
      Left            =   5040
      TabIndex        =   22
      Top             =   5880
      Width           =   375
   End
   Begin VB.Label Label31 
      BackColor       =   &H0000C000&
      Height          =   375
      Left            =   5400
      TabIndex        =   21
      Top             =   5880
      Width           =   3255
   End
   Begin VB.Label Label32 
      BackColor       =   &H0000C000&
      Height          =   375
      Left            =   7080
      TabIndex        =   20
      Top             =   6600
      Width           =   3255
   End
   Begin VB.Label Label33 
      BackColor       =   &H0000C000&
      Height          =   375
      Left            =   5400
      TabIndex        =   19
      Top             =   7200
      Width           =   5295
   End
   Begin VB.Label Label34 
      BackColor       =   &H0000C000&
      Height          =   375
      Left            =   9240
      TabIndex        =   18
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label35 
      BackColor       =   &H0000C000&
      Height          =   2295
      Left            =   10320
      TabIndex        =   17
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label36 
      BackColor       =   &H0000C000&
      Height          =   1575
      Left            =   11160
      TabIndex        =   16
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label37 
      BackColor       =   &H0000C000&
      Height          =   2055
      Left            =   9720
      TabIndex        =   15
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label38 
      BackColor       =   &H0000C000&
      Height          =   2295
      Left            =   11160
      TabIndex        =   14
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label39 
      BackColor       =   &H0000C000&
      Height          =   135
      Left            =   9840
      TabIndex        =   13
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label Label40 
      BackColor       =   &H0000C000&
      Height          =   375
      Left            =   8400
      TabIndex        =   12
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label Label42 
      BackColor       =   &H0000C000&
      Height          =   8295
      Left            =   11880
      TabIndex        =   11
      Top             =   0
      Width           =   135
   End
   Begin VB.Label Label43 
      BackColor       =   &H0000C000&
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   8160
      Width           =   11775
   End
   Begin VB.Label Label44 
      BackColor       =   &H0000C000&
      Height          =   1815
      Left            =   9240
      TabIndex        =   9
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label45 
      BackColor       =   &H0000C000&
      Height          =   375
      Left            =   9120
      TabIndex        =   8
      Top             =   3360
      Width           =   2535
   End
   Begin VB.Label Label46 
      BackColor       =   &H0000C000&
      Height          =   375
      Left            =   9480
      TabIndex        =   7
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label Label47 
      BackColor       =   &H0000C000&
      Height          =   375
      Left            =   8640
      TabIndex        =   6
      Top             =   6600
      Width           =   3255
   End
   Begin VB.Label Label48 
      BackColor       =   &H0000C000&
      Height          =   1575
      Left            =   10440
      TabIndex        =   5
      Top             =   5040
      Width           =   375
   End
   Begin VB.Label Label49 
      BackColor       =   &H0000C000&
      Height          =   855
      Left            =   9000
      TabIndex        =   4
      Top             =   5760
      Width           =   375
   End
   Begin VB.Label Label50 
      BackColor       =   &H0000C000&
      Height          =   855
      Left            =   11160
      TabIndex        =   3
      Top             =   7320
      Width           =   375
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MsgBox "You Win"
Form10.Visible = True
Unload Me
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

Private Sub Label41_Click()
Timer1.Enabled = True
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
Private Sub Label48_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label49_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label50_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Timer1_Timer()
Label41.Caption = Label41.Caption + 1
End Sub

