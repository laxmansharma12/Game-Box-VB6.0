VERSION 5.00
Begin VB.Form Form11 
   Caption         =   "Form11"
   ClientHeight    =   8295
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16530
   LinkTopic       =   "Form11"
   ScaleHeight     =   8295
   ScaleWidth      =   16530
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H80000006&
      Height          =   4335
      Left            =   1200
      TabIndex        =   0
      Top             =   1680
      Width           =   4935
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000B&
         Height          =   1575
         Left            =   1080
         TabIndex        =   4
         Top             =   960
         Width           =   2775
         Begin VB.OptionButton Option1 
            BackColor       =   &H8000000B&
            Caption         =   "Computer"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   0
            TabIndex        =   7
            Top             =   120
            Width           =   1215
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H8000000B&
            Caption         =   "Maths"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   0
            TabIndex        =   6
            Top             =   600
            Width           =   975
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H8000000B&
            Caption         =   "General Knowledge"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   0
            TabIndex        =   5
            Top             =   1080
            Width           =   2175
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "HOME"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1080
         TabIndex        =   3
         Top             =   3240
         Width           =   2775
      End
      Begin VB.CommandButton Command1 
         Caption         =   "PLAY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1080
         TabIndex        =   2
         Top             =   2640
         Width           =   2775
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Enter Your Choice:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   1080
         TabIndex        =   1
         Top             =   600
         Width           =   2775
      End
      Begin VB.Image Image2 
         Height          =   4335
         Left            =   -480
         Picture         =   "quiz-home_Form11.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   11895
      End
   End
   Begin VB.Image Image1 
      Height          =   8880
      Left            =   0
      Picture         =   "quiz-home_Form11.frx":49F1
      Stretch         =   -1  'True
      Top             =   0
      Width           =   16770
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
If Option1.Value = True Then
Form4.Visible = True
Unload Me
ElseIf Option2.Value = True Then
Form5.Visible = True
Unload Me
ElseIf Option3.Value = True Then
Form6.Visible = True
Unload Me
End If
End Sub

Private Sub Command2_Click()
Form1.Visible = True
Unload Me
End Sub

