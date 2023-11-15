VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form4"
   ClientHeight    =   9975
   ClientLeft      =   4755
   ClientTop       =   1050
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   Picture         =   "quiz_Form4_Computer.frx":0000
   ScaleHeight     =   9975
   ScaleWidth      =   11880
   Visible         =   0   'False
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFC0C0&
      Caption         =   "START"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   9240
      Width           =   4200
   End
   Begin VB.Timer Timer1 
      Left            =   10800
      Top             =   120
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000004&
      Caption         =   "Question 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   360
      TabIndex        =   15
      Top             =   4440
      Width           =   11175
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "NEXT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10080
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1080
         Width           =   975
      End
      Begin VB.OptionButton Option7 
         BackColor       =   &H00FF8080&
         Caption         =   "Data Representation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6720
         TabIndex        =   18
         Top             =   960
         Width           =   2895
      End
      Begin VB.OptionButton Option8 
         BackColor       =   &H00FF8080&
         Caption         =   "Flow Chart"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   360
         Left            =   3600
         TabIndex        =   17
         Top             =   960
         Width           =   2895
      End
      Begin VB.OptionButton Option9 
         BackColor       =   &H00FF8080&
         Caption         =   "Data Flow"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   16
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FF8080&
         Caption         =   "3. A pictorial representation of a program or the algorithm is known as _____."
         Height          =   375
         Left            =   240
         TabIndex        =   30
         Top             =   240
         Width           =   10815
      End
      Begin VB.Image Image4 
         Height          =   1575
         Left            =   0
         Picture         =   "quiz_Form4_Computer.frx":643C
         Stretch         =   -1  'True
         Top             =   0
         Width           =   11175
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000004&
      Caption         =   "Question 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   360
      TabIndex        =   13
      Top             =   2880
      Width           =   11175
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FF8080&
         Caption         =   "Assembly Language"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3600
         TabIndex        =   28
         Top             =   840
         Width           =   2895
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00FF8080&
         Caption         =   "Machine Language"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   27
         Top             =   840
         Width           =   2895
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FF8080&
         Caption         =   "High Level Language"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6720
         TabIndex        =   26
         Top             =   840
         Width           =   2895
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "NEXT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10080
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FF8080&
         Caption         =   "2. The first computer were programmed using __________."
         Height          =   375
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Width           =   8295
      End
      Begin VB.Image Image3 
         Height          =   1455
         Left            =   0
         Picture         =   "quiz_Form4_Computer.frx":A9A3
         Stretch         =   -1  'True
         Top             =   0
         Width           =   11175
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000004&
      Caption         =   "Question 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   360
      TabIndex        =   11
      Top             =   1200
      Width           =   11175
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF8080&
         Caption         =   "Data Box Management System"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   495
         Left            =   360
         Picture         =   "quiz_Form4_Computer.frx":EF0A
         TabIndex        =   24
         Top             =   840
         Width           =   2895
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FF8080&
         Caption         =   "Data Base Management System"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   480
         Left            =   3600
         TabIndex        =   23
         Top             =   840
         Width           =   2895
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FF8080&
         Caption         =   "Date Base Management System"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   480
         Left            =   6720
         TabIndex        =   22
         Top             =   840
         Width           =   2895
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "NEXT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10080
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF8080&
         Caption         =   "1. DBMS stands for ___________."
         ForeColor       =   &H80000017&
         Height          =   375
         Left            =   240
         TabIndex        =   25
         Top             =   240
         Width           =   4695
      End
      Begin VB.Image Image2 
         Height          =   1575
         Left            =   0
         Picture         =   "quiz_Form4_Computer.frx":13471
         Stretch         =   -1  'True
         Top             =   0
         Width           =   11175
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H80000004&
      Caption         =   "Question 4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   360
      TabIndex        =   6
      Top             =   6120
      Width           =   11175
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "NEXT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10080
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton Option12 
         BackColor       =   &H00FF8080&
         Caption         =   "Design Page"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3600
         TabIndex        =   9
         Top             =   840
         Width           =   2895
      End
      Begin VB.OptionButton Option11 
         BackColor       =   &H00FF8080&
         Caption         =   "Home Page"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6720
         TabIndex        =   8
         Top             =   840
         Width           =   2895
      End
      Begin VB.OptionButton Option10 
         BackColor       =   &H00FF8080&
         Caption         =   "Screen Page"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   7
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FF8080&
         Caption         =   "4. First page of website is termed as ___________."
         Height          =   375
         Left            =   240
         TabIndex        =   31
         Top             =   240
         Width           =   7095
      End
      Begin VB.Image Image5 
         Height          =   1455
         Left            =   0
         Picture         =   "quiz_Form4_Computer.frx":179D8
         Stretch         =   -1  'True
         Top             =   0
         Width           =   11175
      End
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "END"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9240
      Width           =   4200
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H80000004&
      Caption         =   "Question 5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   360
      TabIndex        =   0
      Top             =   7680
      Width           =   11175
      Begin VB.OptionButton Option15 
         BackColor       =   &H00FF8080&
         Caption         =   "German"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   4
         Top             =   840
         Width           =   2895
      End
      Begin VB.OptionButton Option14 
         BackColor       =   &H00FF8080&
         Caption         =   "Latin"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6720
         TabIndex        =   3
         Top             =   840
         Width           =   2895
      End
      Begin VB.OptionButton Option13 
         BackColor       =   &H00FF8080&
         Caption         =   "Greek"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3600
         TabIndex        =   2
         Top             =   840
         Width           =   2895
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFC0C0&
         Caption         =   "SUBMIT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9840
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FF8080&
         Caption         =   "5. The term 'Computer' is derived from ___________."
         Height          =   375
         Left            =   240
         TabIndex        =   32
         Top             =   240
         Width           =   7455
      End
      Begin VB.Image Image6 
         Height          =   1455
         Left            =   0
         Picture         =   "quiz_Form4_Computer.frx":1BF3F
         Stretch         =   -1  'True
         Top             =   0
         Width           =   11175
      End
   End
   Begin VB.Image Image7 
      Height          =   855
      Left            =   3480
      Picture         =   "quiz_Form4_Computer.frx":204A6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10800
      TabIndex        =   21
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "SCORE BOARD:-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8880
      TabIndex        =   20
      Top             =   720
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   10095
      Left            =   0
      Picture         =   "quiz_Form4_Computer.frx":27E95
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   11895
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim score As Integer

Private Sub Command1_Click()
If Option2.Value = True Then
score = score + 2
Else
score = score + 0
End If
Frame1.Visible = False
Frame2.Visible = True
Command5.Visible = False
End Sub

Private Sub Command2_Click()
If Option6.Value = True Then
score = score + 2
Else
score = score + 0
End If
Frame2.Visible = False
Frame3.Visible = True
Command5.Visible = False
End Sub

Private Sub Command3_Click()
If Option8.Value = True Then
score = score + 2
Else
score = score + 0
End If
Frame3.Visible = False
Frame4.Visible = True
Command5.Visible = False
End Sub


Private Sub Command4_Click()
If Option11.Value = True Then
score = score + 2
Else
score = score + 0
End If
Frame4.Visible = False
Frame5.Visible = True
Command5.Visible = False
End Sub

Private Sub Command5_Click()
Form11.Visible = True
Unload Me
End Sub

Private Sub Command6_Click()
If Option14.Value = True Then
score = score + 2
Else
score = score + 0
End If
Frame1.Visible = True
Frame2.Visible = True
Frame3.Visible = True
Frame4.Visible = True
Command5.Visible = True

If Option1.Value = True Then
Option1.BackColor = vbRed
ElseIf Option3.Value = True Then
Option3.BackColor = vbRed
End If

If Option5.Value = True Then
Option5.BackColor = vbRed
ElseIf Option4.Value = True Then
Option4.BackColor = vbRed
End If

If Option7.Value = True Then
Option7.BackColor = vbRed
ElseIf Option9.Value = True Then
Option9.BackColor = vbRed
End If

If Option10.Value = True Then
Option10.BackColor = vbRed
ElseIf Option12.Value = True Then
Option12.BackColor = vbRed
End If

If Option15.Value = True Then
Option15.BackColor = vbRed
ElseIf Option13.Value = True Then
Option13.BackColor = vbRed
End If

Option2.BackColor = vbGreen
Option6.BackColor = vbGreen
Option8.BackColor = vbGreen
Option11.BackColor = vbGreen
Option14.BackColor = vbGreen
Label3.Caption = score
End Sub

Private Sub Form_Load()
 score = 0
Frame1.Visible = True
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = False
Command5.Visible = False
End Sub

