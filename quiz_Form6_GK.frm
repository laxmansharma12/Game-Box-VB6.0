VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H8000000E&
   Caption         =   "Form6"
   ClientHeight    =   9750
   ClientLeft      =   5130
   ClientTop       =   1050
   ClientWidth     =   11655
   LinkTopic       =   "Form6"
   ScaleHeight     =   9750
   ScaleWidth      =   11655
   Visible         =   0   'False
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
      Height          =   1335
      Left            =   240
      TabIndex        =   25
      Top             =   4560
      Width           =   11175
      Begin VB.OptionButton Option7 
         BackColor       =   &H00FF8080&
         Caption         =   "Lucknow"
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
         Left            =   480
         TabIndex        =   29
         Top             =   720
         Width           =   2895
      End
      Begin VB.OptionButton Option9 
         BackColor       =   &H00FF8080&
         Caption         =   "Hyderabad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   3600
         TabIndex        =   28
         Top             =   720
         Width           =   2895
      End
      Begin VB.OptionButton Option8 
         BackColor       =   &H00FF8080&
         Caption         =   "Allahabad"
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
         TabIndex        =   27
         Top             =   720
         Width           =   3015
      End
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
         TabIndex        =   26
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FF8080&
         Caption         =   "3. Where is the ""National Academy of Science"" situated ?"
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
         Width           =   8055
      End
      Begin VB.Image Image5 
         Height          =   1455
         Left            =   0
         Picture         =   "quiz_Form6_GK.frx":0000
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
      TabIndex        =   19
      Top             =   3000
      Width           =   11175
      Begin VB.OptionButton Option6 
         BackColor       =   &H00FF8080&
         Caption         =   "1975"
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
         TabIndex        =   23
         Top             =   840
         Width           =   2895
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FF8080&
         Caption         =   "1973"
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
         Left            =   480
         TabIndex        =   22
         Top             =   840
         Width           =   2895
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FF8080&
         Caption         =   "1974"
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
         TabIndex        =   21
         Top             =   840
         Width           =   3015
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
         TabIndex        =   20
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FF8080&
         Caption         =   "2. In which year was internal emergency imposed in India ?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   24
         Top             =   240
         Width           =   8295
      End
      Begin VB.Image Image2 
         Height          =   1455
         Left            =   0
         Picture         =   "quiz_Form6_GK.frx":4567
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
      Height          =   1455
      Left            =   240
      TabIndex        =   13
      Top             =   1440
      Width           =   11175
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
         TabIndex        =   17
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FF8080&
         Caption         =   "Brazil"
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
         TabIndex        =   16
         Top             =   840
         Width           =   2895
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FF8080&
         Caption         =   "Cuba"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   480
         TabIndex        =   15
         Top             =   840
         Width           =   2895
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF8080&
         Caption         =   "India"
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
         TabIndex        =   14
         Top             =   840
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF8080&
         Caption         =   "1. Which country is known as the""Sugar Bowl of the World"" ?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   10335
      End
      Begin VB.Image Image6 
         Height          =   1695
         Left            =   0
         Picture         =   "quiz_Form6_GK.frx":8ACE
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
      Height          =   1335
      Left            =   240
      TabIndex        =   7
      Top             =   6000
      Width           =   11175
      Begin VB.OptionButton Option10 
         BackColor       =   &H00FF8080&
         Caption         =   "White"
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
         TabIndex        =   11
         Top             =   720
         Width           =   2895
      End
      Begin VB.OptionButton Option11 
         BackColor       =   &H00FF8080&
         Caption         =   "Green"
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
         TabIndex        =   10
         Top             =   720
         Width           =   3015
      End
      Begin VB.OptionButton Option12 
         BackColor       =   &H00FF8080&
         Caption         =   "Saffron"
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
         Left            =   480
         TabIndex        =   9
         Top             =   720
         Width           =   2895
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
         Height          =   360
         Left            =   10080
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FF8080&
         Caption         =   "4. Which colour symbolises peace ?"
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
         TabIndex        =   12
         Top             =   240
         Width           =   5175
      End
      Begin VB.Image Image3 
         Height          =   1455
         Left            =   0
         Picture         =   "quiz_Form6_GK.frx":D035
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
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9000
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
      Left            =   240
      TabIndex        =   0
      Top             =   7440
      Width           =   11175
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
         Height          =   360
         Left            =   9840
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton Option13 
         BackColor       =   &H00FF8080&
         Caption         =   "Jan 11"
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
         Left            =   480
         TabIndex        =   3
         Top             =   840
         Width           =   2895
      End
      Begin VB.OptionButton Option14 
         BackColor       =   &H00FF8080&
         Caption         =   "Jan 13"
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
         TabIndex        =   2
         Top             =   840
         Width           =   3015
      End
      Begin VB.OptionButton Option15 
         BackColor       =   &H00FF8080&
         Caption         =   "Jan 12"
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
         TabIndex        =   1
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FF8080&
         Caption         =   "5. National Youth Day is celebrated on ___________."
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
         TabIndex        =   5
         Top             =   240
         Width           =   7455
      End
      Begin VB.Image Image4 
         Height          =   1455
         Left            =   0
         Picture         =   "quiz_Form6_GK.frx":1159C
         Stretch         =   -1  'True
         Top             =   0
         Width           =   11175
      End
   End
   Begin VB.Image Image7 
      Height          =   735
      Left            =   2520
      Picture         =   "quiz_Form6_GK.frx":15B03
      Stretch         =   -1  'True
      Top             =   120
      Width           =   6255
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
      Left            =   8760
      TabIndex        =   32
      Top             =   960
      Width           =   1935
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
      TabIndex        =   31
      Top             =   960
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   10095
      Left            =   0
      Picture         =   "quiz_Form6_GK.frx":1F62F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "Form6"
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
If Option10.Value = True Then
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
If Option15.Value = True Then
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

If Option12.Value = True Then
Option12.BackColor = vbRed
ElseIf Option11.Value = True Then
Option11.BackColor = vbRed
End If

If Option13.Value = True Then
Option13.BackColor = vbRed
ElseIf Option14.Value = True Then
Option14.BackColor = vbRed
End If

Option2.BackColor = vbGreen
Option6.BackColor = vbGreen
Option8.BackColor = vbGreen
Option10.BackColor = vbGreen
Option15.BackColor = vbGreen
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


