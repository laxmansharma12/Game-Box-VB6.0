VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H80000002&
   Caption         =   "Form2"
   ClientHeight    =   8880
   ClientLeft      =   3975
   ClientTop       =   1620
   ClientWidth     =   14535
   LinkTopic       =   "Form2"
   ScaleHeight     =   8880
   ScaleWidth      =   14535
   Visible         =   0   'False
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   17.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5760
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "NEW CARD"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   17.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4560
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "START"
      DownPicture     =   "blackJack_Form2.frx":0000
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   17.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3360
      Width           =   3135
   End
   Begin VB.Image Image14 
      Height          =   4335
      Left            =   11760
      Picture         =   "blackJack_Form2.frx":1BD0
      Stretch         =   -1  'True
      Top             =   4080
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Image Image13 
      Height          =   4335
      Left            =   11280
      Picture         =   "blackJack_Form2.frx":5E539
      Stretch         =   -1  'True
      Top             =   3600
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Image Image12 
      Height          =   4335
      Left            =   10800
      Picture         =   "blackJack_Form2.frx":B7C37
      Stretch         =   -1  'True
      Top             =   3120
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Image Image11 
      Height          =   4335
      Left            =   10320
      Picture         =   "blackJack_Form2.frx":115547
      Stretch         =   -1  'True
      Top             =   2640
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Image Image10 
      Height          =   4335
      Left            =   9840
      Picture         =   "blackJack_Form2.frx":136D7E
      Stretch         =   -1  'True
      Top             =   2160
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Image Image9 
      Height          =   4335
      Left            =   9360
      Picture         =   "blackJack_Form2.frx":155A94
      Stretch         =   -1  'True
      Top             =   1680
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Image Image8 
      Height          =   4335
      Left            =   8880
      Picture         =   "blackJack_Form2.frx":173095
      Stretch         =   -1  'True
      Top             =   1200
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   735
      Left            =   7680
      TabIndex        =   8
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      Caption         =   "TOTAL SUM:"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   15
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   5280
      TabIndex        =   7
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      Caption         =   "BLACK JACK"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   855
      Left            =   4920
      TabIndex        =   6
      Top             =   240
      Width           =   3855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   735
      Left            =   7680
      TabIndex        =   5
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   735
      Left            =   6480
      TabIndex        =   4
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   735
      Left            =   5400
      TabIndex        =   3
      Top             =   1200
      Width           =   735
   End
   Begin VB.Image Image7 
      Height          =   4335
      Left            =   2400
      Picture         =   "blackJack_Form2.frx":18EAE8
      Stretch         =   -1  'True
      Top             =   3480
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Image Image6 
      Height          =   4335
      Left            =   1920
      Picture         =   "blackJack_Form2.frx":1A99CB
      Stretch         =   -1  'True
      Top             =   3000
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Image Image5 
      Height          =   4335
      Left            =   1440
      Picture         =   "blackJack_Form2.frx":1C3522
      Stretch         =   -1  'True
      Top             =   2520
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Image Image4 
      Height          =   4335
      Left            =   960
      Picture         =   "blackJack_Form2.frx":1DC70B
      Stretch         =   -1  'True
      Top             =   2040
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Image Image3 
      Height          =   4335
      Left            =   480
      Picture         =   "blackJack_Form2.frx":1F4141
      Stretch         =   -1  'True
      Top             =   1560
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Image Image2 
      Height          =   4335
      Left            =   0
      Picture         =   "blackJack_Form2.frx":20997A
      Stretch         =   -1  'True
      Top             =   1080
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   9015
      Left            =   -240
      Picture         =   "blackJack_Form2.frx":238BCF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14775
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Randomize
Label1.Caption = Int(13 * Rnd) + 1
Randomize
Label2.Caption = Int(13 * Rnd) + 1

If (Val(Label1.Caption) = 1) Then
Image2.Visible = True
ElseIf (Val(Label1.Caption) = 2) Then
Image3.Visible = True
ElseIf (Val(Label1.Caption) = 3) Then
Image4.Visible = True
ElseIf (Val(Label1.Caption) = 4) Then
Image5.Visible = True
ElseIf (Val(Label1.Caption) = 5) Then
Image6.Visible = True
ElseIf (Val(Label1.Caption) = 6) Then
Image7.Visible = True
ElseIf (Val(Label1.Caption) = 7) Then
Image8.Visible = True
ElseIf (Val(Label1.Caption) = 8) Then
Image9.Visible = True
ElseIf (Val(Label1.Caption) = 9) Then
Image10.Visible = True
ElseIf (Val(Label1.Caption) = 10) Then
Image11.Visible = True
ElseIf (Val(Label1.Caption) = 11) Then
Image12.Visible = True
ElseIf (Val(Label1.Caption) = 12) Then
Image13.Visible = True
ElseIf (Val(Label1.Caption) = 13) Then
Image14.Visible = True
End If

If (Val(Label2.Caption) = 1) Then
Image2.Visible = True
ElseIf (Val(Label2.Caption) = 2) Then
Image3.Visible = True
ElseIf (Val(Label2.Caption) = 3) Then
Image4.Visible = True
ElseIf (Val(Label2.Caption) = 4) Then
Image5.Visible = True
ElseIf (Val(Label2.Caption) = 5) Then
Image6.Visible = True
ElseIf (Val(Label2.Caption) = 6) Then
Image7.Visible = True
ElseIf (Val(Label2.Caption) = 7) Then
Image8.Visible = True
ElseIf (Val(Label2.Caption) = 8) Then
Image9.Visible = True
ElseIf (Val(Label2.Caption) = 9) Then
Image10.Visible = True
ElseIf (Val(Label2.Caption) = 10) Then
Image11.Visible = True
ElseIf (Val(Label2.Caption) = 11) Then
Image12.Visible = True
ElseIf (Val(Label2.Caption) = 12) Then
Image13.Visible = True
ElseIf (Val(Label2.Caption) = 13) Then
Image14.Visible = True
End If

Command1.Enabled = False

If (Val(Label1.Caption) > 10) Then
Label1.Caption = 10
End If

If (Val(Label1.Caption) = 1) Then
Label1.Caption = 11
End If

If (Val(Label2.Caption) = 1) Then
Label2.Caption = 11
End If

If (Val(Label2.Caption) > 10) Then
Label2.Caption = 10
End If

Label6.Caption = Val(Label1.Caption) + Val(Label2.Caption)

If (Val(Label6.Caption) < 21) Then
Command2.Enabled = True
MsgBox "To draw a new card,click on new card button", 64 + 0, "Information"
ElseIf (Val(Label6.Caption) = 21) Then
MsgBox "Woohooo! you've got Black Jack", 64 + 0, "Congratulations?????????"
Else
MsgBox "You are out of game. Click on Exit button", 16 + 0, "Oops!"
End If
End Sub

Private Sub Command2_Click()
Randomize
Label3.Caption = Int(Rnd() * 13) + 1

If (Val(Label3.Caption) = 1) Then
Label3.Caption = 11
End If

If (Val(Label3.Caption) > 10) Then
Label3.Caption = 10
End If

Label6.Caption = Val(Label6.Caption) + Val(Label3.Caption)
If (Val(Label3.Caption) = 1) Then
Image2.Visible = True
ElseIf (Val(Label3.Caption) = 2) Then
Image3.Visible = True
ElseIf (Val(Label3.Caption) = 3) Then
Image4.Visible = True
ElseIf (Val(Label3.Caption) = 4) Then
Image5.Visible = True
ElseIf (Val(Label3.Caption) = 5) Then
Image6.Visible = True
ElseIf (Val(Label3.Caption) = 6) Then
Image7.Visible = True
ElseIf (Val(Label3.Caption) = 7) Then
Image8.Visible = True
ElseIf (Val(Label3.Caption) = 8) Then
Image9.Visible = True
ElseIf (Val(Label3.Caption) = 9) Then
Image10.Visible = True
ElseIf (Val(Label3.Caption) = 10) Then
Image11.Visible = True
ElseIf (Val(Label3.Caption) = 11) Then
Image12.Visible = True
ElseIf (Val(Label3.Caption) = 12) Then
Image13.Visible = True
ElseIf (Val(Label3.Caption) = 13) Then
Image14.Visible = True
End If
If (Val(Label6.Caption) = 21) Then
MsgBox "Woohooo! you've got Black Jack", 64 + 0, "Congratulations!!!!!"
Command2.Enabled = False
ElseIf (Val(Label6.Caption) > 21) Then
MsgBox "You are out of game. Click on Exit button", 16 + 0, "Oops!"
Command2.Enabled = False
End If
End Sub

Private Sub Command3_Click()
Form1.Show
Unload Me
End Sub

