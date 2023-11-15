VERSION 5.00
Begin VB.Form Form10 
   Caption         =   "MAZE"
   ClientHeight    =   8295
   ClientLeft      =   -7395
   ClientTop       =   -2235
   ClientWidth     =   16530
   LinkTopic       =   "Form10"
   ScaleHeight     =   8295
   ScaleWidth      =   16530
   Visible         =   0   'False
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
      Left            =   6960
      TabIndex        =   6
      Top             =   6720
      Width           =   3015
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
      Left            =   6960
      TabIndex        =   5
      Top             =   6120
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   2775
      Left            =   6960
      TabIndex        =   0
      Top             =   3240
      Width           =   3015
      Begin VB.OptionButton Option3 
         BackColor       =   &H000000C0&
         Caption         =   "Hard"
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
         Left            =   360
         TabIndex        =   4
         Top             =   2040
         Width           =   2295
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H0000C000&
         Caption         =   "Medium"
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
         Left            =   360
         TabIndex        =   3
         Top             =   1440
         Width           =   2295
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000D&
         Caption         =   "Easy"
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
         Left            =   360
         TabIndex        =   2
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         Caption         =   "Select difficulty:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   495
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Image Image2 
      Height          =   2055
      Left            =   7080
      Picture         =   "maze-home_Form10.frx":0000
      Stretch         =   -1  'True
      Top             =   600
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   8880
      Left            =   0
      Picture         =   "maze-home_Form10.frx":1C4B5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   16770
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Option1.Value = True Then
Form12.Visible = True
Unload Me
ElseIf Option2.Value = True Then
Form3.Visible = True
Unload Me
ElseIf Option3.Value = True Then
Form9.Visible = True
Unload Me
End If
End Sub

Private Sub Command2_Click()
Form1.Visible = True
Unload Me
End Sub

