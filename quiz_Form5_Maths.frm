VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H8000000E&
   Caption         =   "Form5"
   ClientHeight    =   9615
   ClientLeft      =   4935
   ClientTop       =   1230
   ClientWidth     =   11535
   LinkTopic       =   "Form5"
   Picture         =   "quiz_Form5_Maths.frx":0000
   ScaleHeight     =   9615
   ScaleWidth      =   11535
   Visible         =   0   'False
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
      Height          =   1455
      Left            =   240
      TabIndex        =   21
      Top             =   1200
      Width           =   11055
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF8080&
         Caption         =   "3  Months"
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
         TabIndex        =   25
         Top             =   840
         Width           =   2775
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FF8080&
         Caption         =   "4  Months"
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
         Left            =   6960
         TabIndex        =   24
         Top             =   840
         Width           =   2775
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FF8080&
         Caption         =   "6  Months"
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
         Left            =   3720
         TabIndex        =   23
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
         Left            =   9960
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF8080&
         Caption         =   "1. How many months have 120 Days ?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Width           =   5415
      End
      Begin VB.Image Image6 
         Height          =   1455
         Left            =   0
         Picture         =   "quiz_Form5_Maths.frx":643C
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
      Left            =   240
      TabIndex        =   16
      Top             =   2760
      Width           =   11055
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
         Left            =   9960
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FF8080&
         Caption         =   "Square of a circle"
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
         Left            =   7080
         TabIndex        =   19
         Top             =   840
         Width           =   2655
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FF8080&
         Caption         =   "Area of a circle"
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
         TabIndex        =   18
         Top             =   840
         Width           =   2775
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00FF8080&
         Caption         =   "Perimeter of a circle"
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
         Left            =   3720
         TabIndex        =   17
         Top             =   840
         Width           =   3015
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FF8080&
         Caption         =   "2. The circumference of the circle is also sometimes called __________."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   30
         Top             =   240
         Width           =   10095
      End
      Begin VB.Image Image2 
         Height          =   1455
         Left            =   0
         Picture         =   "quiz_Form5_Maths.frx":A9A3
         Stretch         =   -1  'True
         Top             =   0
         Width           =   11175
      End
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
      Height          =   1455
      Left            =   240
      TabIndex        =   11
      Top             =   4320
      Width           =   11055
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
         Left            =   9960
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton Option8 
         BackColor       =   &H00FF8080&
         Caption         =   "Length * Breadth"
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
         TabIndex        =   14
         Top             =   840
         Width           =   2895
      End
      Begin VB.OptionButton Option9 
         BackColor       =   &H00FF8080&
         Caption         =   "Length * Breadth * Height"
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
         Left            =   3720
         TabIndex        =   13
         Top             =   840
         Width           =   3015
      End
      Begin VB.OptionButton Option7 
         BackColor       =   &H00FF8080&
         Caption         =   "2(Length * Breadth)"
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
         Left            =   7080
         TabIndex        =   12
         Top             =   840
         Width           =   2655
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FF8080&
         Caption         =   "3. The area of rectangle is equal to _________."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   240
         Width           =   6615
      End
      Begin VB.Image Image3 
         Height          =   1455
         Left            =   0
         Picture         =   "quiz_Form5_Maths.frx":EF0A
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
      Left            =   240
      TabIndex        =   6
      Top             =   5880
      Width           =   11055
      Begin VB.OptionButton Option10 
         BackColor       =   &H00FF8080&
         Caption         =   "26"
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
         TabIndex        =   10
         Top             =   840
         Width           =   2895
      End
      Begin VB.OptionButton Option11 
         BackColor       =   &H00FF8080&
         Caption         =   "25"
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
         Left            =   3720
         TabIndex        =   9
         Top             =   840
         Width           =   3015
      End
      Begin VB.OptionButton Option12 
         BackColor       =   &H00FF8080&
         Caption         =   "24"
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
         Left            =   7080
         TabIndex        =   8
         Top             =   840
         Width           =   2535
      End
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
         Left            =   9960
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FF8080&
         Caption         =   "4. Solve 23 + 3 % 3 = _________."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   31
         Top             =   240
         Width           =   4695
      End
      Begin VB.Image Image4 
         Height          =   1455
         Left            =   0
         Picture         =   "quiz_Form5_Maths.frx":13471
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
      Height          =   495
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9000
      Width           =   4080
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
      Left            =   240
      TabIndex        =   0
      Top             =   7440
      Width           =   11055
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
         TabIndex        =   4
         Top             =   960
         Width           =   1095
      End
      Begin VB.OptionButton Option13 
         BackColor       =   &H00FF8080&
         Caption         =   "8"
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
         Left            =   7080
         TabIndex        =   3
         Top             =   840
         Width           =   2535
      End
      Begin VB.OptionButton Option14 
         BackColor       =   &H00FF8080&
         Caption         =   "6"
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
         Left            =   3720
         TabIndex        =   2
         Top             =   840
         Width           =   3015
      End
      Begin VB.OptionButton Option15 
         BackColor       =   &H00FF8080&
         Caption         =   "4"
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
         TabIndex        =   1
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FF8080&
         Caption         =   "5. How many vertices are present on a cube ?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   32
         Top             =   240
         Width           =   6495
      End
      Begin VB.Image Image5 
         Height          =   1455
         Left            =   0
         Picture         =   "quiz_Form5_Maths.frx":179D8
         Stretch         =   -1  'True
         Top             =   0
         Width           =   11175
      End
   End
   Begin VB.Image Image7 
      Height          =   855
      Left            =   3120
      Picture         =   "quiz_Form5_Maths.frx":1BF3F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4815
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
      TabIndex        =   27
      Top             =   720
      Width           =   1815
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
      Left            =   10680
      TabIndex        =   26
      Top             =   720
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   10935
      Left            =   -120
      Picture         =   "quiz_Form5_Maths.frx":22E37
      Stretch         =   -1  'True
      Top             =   -360
      Width           =   12615
   End
End
Attribute VB_Name = "Form5"
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
If Option12.Value = True Then
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
If Option13.Value = True Then
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

If Option11.Value = True Then
Option11.BackColor = vbRed
ElseIf Option10.Value = True Then
Option10.BackColor = vbRed
End If

If Option14.Value = True Then
Option14.BackColor = vbRed
ElseIf Option15.Value = True Then
Option15.BackColor = vbRed
End If

Option2.BackColor = vbGreen
Option6.BackColor = vbGreen
Option8.BackColor = vbGreen
Option12.BackColor = vbGreen
Option13.BackColor = vbGreen
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

