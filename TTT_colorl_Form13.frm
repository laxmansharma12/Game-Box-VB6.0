VERSION 5.00
Begin VB.Form Form13 
   BackColor       =   &H00FF0000&
   Caption         =   "Form13"
   ClientHeight    =   5505
   ClientLeft      =   7830
   ClientTop       =   2970
   ClientWidth     =   6210
   LinkTopic       =   "Form13"
   ScaleHeight     =   5505
   ScaleWidth      =   6210
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   4
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   5
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   6
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   7
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   8
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "New Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   0
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "TIC TAC TOE"
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
      Left            =   1800
      TabIndex        =   15
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "'s Turn"
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
      Left            =   840
      TabIndex        =   14
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "'s Turn"
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
      Left            =   5160
      TabIndex        =   13
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
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
      Left            =   240
      TabIndex        =   12
      Top             =   720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
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
      Left            =   4560
      TabIndex        =   11
      Top             =   720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   5160
      TabIndex        =   10
      Top             =   5040
      Width           =   1095
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim USER As Byte
Dim winner1 As Boolean

Private Sub Command1_Click(Index As Integer)
If USER = 0 Then
Command1(Index).BackColor = vbGreen
Label5.Visible = True
Label4.Visible = False
USER = 1
Else
Command1(Index).BackColor = vbRed
USER = 0
Label4.Visible = True
Label5.Visible = False
End If
Command1(Index).Enabled = False
If Command1(0).BackColor = vbGreen And Command1(1).BackColor = vbGreen And Command1(2).BackColor = vbGreen Then
MsgBox "USER GREEN : WIN"
winner1 = True
Command1(3).Enabled = False
Command1(4).Enabled = False
Command1(5).Enabled = False
Command1(6).Enabled = False
Command1(7).Enabled = True
Command1(8).Enabled = False
End If

If Command1(3).BackColor = vbGreen And Command1(4).BackColor = vbGreen And Command1(5).BackColor = vbGreen Then
MsgBox "USER GREEN : WIN"
winner1 = True
Command1(0).Enabled = False
Command1(1).Enabled = False
Command1(2).Enabled = False
Command1(6).Enabled = False
Command1(7).Enabled = True
Command1(8).Enabled = False
End If

If Command1(6).BackColor = vbGreen And Command1(7).BackColor = vbGreen And Command1(8).BackColor = vbGreen Then
MsgBox "USER GREEN : WIN"
winner1 = True
Command1(0).Enabled = False
Command1(1).Enabled = False
Command1(2).Enabled = False
Command1(3).Enabled = False
Command1(4).Enabled = True
Command1(5).Enabled = False
End If

If Command1(0).BackColor = vbGreen And Command1(3).BackColor = vbGreen And Command1(6).BackColor = vbGreen Then
MsgBox "USER GREEN : WIN"
winner1 = True
Command1(1).Enabled = False
Command1(2).Enabled = False
Command1(4).Enabled = False
Command1(5).Enabled = False
Command1(7).Enabled = True
Command1(8).Enabled = False
End If

If Command1(1).BackColor = vbGreen And Command1(4).BackColor = vbGreen And Command1(7).BackColor = vbGreen Then
MsgBox "USER GREEN : WIN"
winner1 = True
Command1(0).Enabled = False
Command1(2).Enabled = False
Command1(3).Enabled = False
Command1(5).Enabled = True
Command1(6).Enabled = False
Command1(8).Enabled = False
End If

If Command1(2).BackColor = vbGreen And Command1(5).BackColor = vbGreen And Command1(8).BackColor = vbGreen Then
MsgBox "USER GREEN : WIN"
winner1 = True
Command1(0).Enabled = False
Command1(1).Enabled = False
Command1(3).Enabled = True
Command1(4).Enabled = False
Command1(6).Enabled = False
Command1(7).Enabled = False
End If

If Command1(0).BackColor = vbGreen And Command1(4).BackColor = vbGreen And Command1(8).BackColor = vbGreen Then
MsgBox "USER GREEN : WIN"
winner1 = True
Command1(1).Enabled = False
Command1(2).Enabled = False
Command1(3).Enabled = False
Command1(5).Enabled = True
Command1(6).Enabled = False
Command1(7).Enabled = False
End If

If Command1(2).BackColor = vbGreen And Command1(4).BackColor = vbGreen And Command1(6).BackColor = vbGreen Then
MsgBox "USER GREEN : WIN"
winner1 = True
Command1(0).Enabled = False
Command1(1).Enabled = False
Command1(3).Enabled = False
Command1(5).Enabled = False
Command1(7).Enabled = True
Command1(8).Enabled = False
End If

'-----------------------------------------------------------------------------

If Command1(0).BackColor = vbRed And Command1(1).BackColor = vbRed And Command1(2).BackColor = vbRed Then
MsgBox "USER RED : WIN"
winner1 = True
Command1(3).Enabled = False
Command1(4).Enabled = False
Command1(5).Enabled = False
Command1(6).Enabled = False
Command1(7).Enabled = False
Command1(8).Enabled = True
End If

If Command1(3).BackColor = vbRed And Command1(4).BackColor = vbRed And Command1(5).BackColor = vbRed Then
MsgBox "USER RED : WIN"
winner1 = True
Command1(0).Enabled = False
Command1(1).Enabled = False
Command1(2).Enabled = False
Command1(6).Enabled = False
Command1(7).Enabled = False
Command1(8).Enabled = True
End If

If Command1(6).BackColor = vbRed And Command1(7).BackColor = vbRed And Command1(8).BackColor = vbRed Then
MsgBox "USER RED : WIN"
winner1 = True
Command1(0).Enabled = False
Command1(1).Enabled = False
Command1(2).Enabled = False
Command1(3).Enabled = False
Command1(4).Enabled = False
Command1(5).Enabled = True
End If

If Command1(0).BackColor = vbRed And Command1(3).BackColor = vbRed And Command1(6).BackColor = vbRed Then
MsgBox "USER RED : WIN"
winner1 = True
Command1(1).Enabled = False
Command1(2).Enabled = False
Command1(4).Enabled = False
Command1(5).Enabled = False
Command1(7).Enabled = False
Command1(8).Enabled = True
End If

If Command1(1).BackColor = vbRed And Command1(4).BackColor = vbRed And Command1(7).BackColor = vbRed Then
MsgBox "USER RED : WIN"
winner1 = True
Command1(0).Enabled = False
Command1(2).Enabled = False
Command1(3).Enabled = False
Command1(5).Enabled = False
Command1(6).Enabled = False
Command1(8).Enabled = True
End If

If Command1(2).BackColor = vbRed And Command1(5).BackColor = vbRed And Command1(8).BackColor = vbRed Then
MsgBox "USER RED : WIN"
winner1 = True
Command1(0).Enabled = False
Command1(1).Enabled = False
Command1(3).Enabled = False
Command1(4).Enabled = False
Command1(6).Enabled = False
Command1(7).Enabled = True
End If

If Command1(0).BackColor = vbRed And Command1(4).BackColor = vbRed And Command1(8).BackColor = vbRed Then
MsgBox "USER RED : WIN"
winner1 = True
Command1(1).Enabled = False
Command1(2).Enabled = False
Command1(3).Enabled = False
Command1(5).Enabled = False
Command1(6).Enabled = False
Command1(7).Enabled = True
End If

If Command1(2).BackColor = vbRed And Command1(4).BackColor = vbRed And Command1(6).BackColor = vbRed Then
MsgBox "USER RED : WIN"
winner1 = True
Command1(0).Enabled = False
Command1(1).Enabled = False
Command1(3).Enabled = False
Command1(5).Enabled = False
Command1(7).Enabled = False
Command1(8).Enabled = True
End If
If Command1(0).Enabled = False And Command1(1).Enabled = False And Command1(2).Enabled = False And Command1(3).Enabled = False And Command1(4).Enabled = False And Command1(5).Enabled = False And Command1(6).Enabled = False And Command1(7).Enabled = False And Command1(8).Enabled = False Then
MsgBox "No Winner. Click on New Game", vbCritical + vbOKOnly, "Oops!!"
End If
End Sub

Private Sub Command2_Click()
Command1(0).BackColor = &H8000000F
Command1(1).BackColor = &H8000000F
Command1(2).BackColor = &H8000000F
Command1(3).BackColor = &H8000000F
Command1(4).BackColor = &H8000000F
Command1(5).BackColor = &H8000000F
Command1(6).BackColor = &H8000000F
Command1(7).BackColor = &H8000000F
Command1(8).BackColor = &H8000000F
Command1(0).Enabled = True
Command1(1).Enabled = True
Command1(2).Enabled = True
Command1(3).Enabled = True
Command1(4).Enabled = True
Command1(5).Enabled = True
Command1(6).Enabled = True
Command1(7).Enabled = True
Command1(8).Enabled = True
End Sub

Private Sub Form_Load()
USER = 0
winner1 = False
Label5.Visible = False
Label4.Visible = True
End Sub

Private Sub Label6_Click()
Load Form1
Form1.Show
Unload Me
End Sub

