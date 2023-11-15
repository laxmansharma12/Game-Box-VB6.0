VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form8 
   Caption         =   "Form8"
   ClientHeight    =   8325
   ClientLeft      =   4365
   ClientTop       =   1230
   ClientWidth     =   13080
   LinkTopic       =   "Form8"
   Picture         =   "sign-up_Form8.frx":0000
   ScaleHeight     =   8325
   ScaleWidth      =   13080
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   5040
      Top             =   9120
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Ruby Kumari\Pictures\LOGIN-DATABSE.MDB;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Ruby Kumari\Pictures\LOGIN-DATABSE.MDB;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "tableLogin"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   7455
      Left            =   3840
      TabIndex        =   0
      Top             =   480
      Width           =   5655
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFF80&
         Caption         =   "Back"
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
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   6480
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Register"
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
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   6480
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2160
         TabIndex        =   4
         Top             =   3840
         Width           =   3375
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2160
         TabIndex        =   3
         Top             =   4440
         Width           =   2535
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2160
         TabIndex        =   2
         Top             =   5040
         Width           =   2055
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2160
         TabIndex        =   1
         Top             =   5640
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "REGISTERATION FORM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         TabIndex        =   11
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Name:"
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
         Left            =   1200
         TabIndex        =   10
         Top             =   3840
         Width           =   855
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Player ID:"
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
         TabIndex        =   9
         Top             =   4440
         Width           =   1335
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Password:"
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
         Left            =   840
         TabIndex        =   8
         Top             =   5040
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Conf. Password:"
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
         Left            =   120
         TabIndex        =   7
         Top             =   5640
         Width           =   2055
      End
      Begin VB.Image Image2 
         Height          =   1815
         Left            =   1920
         Picture         =   "sign-up_Form8.frx":215A
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   1935
      End
   End
   Begin VB.Image Image1 
      Height          =   8400
      Left            =   0
      Picture         =   "sign-up_Form8.frx":80F0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13080
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim conn As New ADODB.Connection

Private Sub Command1_Click()
Load Form1
Form1.Show
rs.Close
conn.Close
Unload Me
End Sub

Private Sub Command2_Click()
If Text1.Text = "" Then
MsgBox "Enter Your Name", vbInformation + vbOKOnly, "sign-up"
Text1.SetFocus
Exit Sub
End If

If Text5.Text = "" Then
MsgBox "Enter Your playerID", vbInformation + vbOKOnly, "sign-up"
Text5.SetFocus
Exit Sub
End If

If Text6.Text = "" Then
MsgBox "Enter your Password", vbInformation + vbOKOnly, "sign-up"
Text6.SetFocus
Exit Sub
End If

If Text7.Text = "" Then
MsgBox "Please confirm your password", vbInformation + vbOKOnly, "sign-up"
Text7.SetFocus
Exit Sub
End If

If Text1.Text <> "" And Text5.Text <> "" And Text6.Text <> "" And Text7.Text <> "" Then
If rs.BOF = False Then
rs.MoveLast
End If

rs.AddNew
rs("Name") = Text1.Text
rs("playerID") = Text5.Text
If Text6.Text = Text7.Text Then
rs("Password") = Text7.Text
Else
MsgBox "Oops!! password miss match. Enter again"
Text6.SetFocus
Text6.Text = ""
Text7.Text = ""
End If
MsgBox "Congrats!!1 You've successfull Registered", vbInformation + vbOKOnly, "login"
End If
rs.update
Text1.Text = ""
Text6.Text = ""
Text5.Text = ""
Text7.Text = ""
End Sub

Private Sub Form_Load()
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Ruby Kumari\Pictures\LOGIN-DATABSE.MDB;Persist Security Info=False"
'conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\home\Documents\vb project\LOGIN-DATABSE.MDB;Persist Security Info=False"
rs.Open "select * from tableLogin", conn, adOpenDynamic, adLockOptimistic, adCmdText
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text1.Text <> "" Then
Text5.SetFocus
ElseIf (KeyAscii < 65 And KeyAscii <> 8 And KeyAscii > 32) Or (KeyAscii > 90 And KeyAscii < 97) Or (KeyAscii > 122) Then
KeyAscii = 0
MsgBox "Enter characters only"
End If
End Sub


Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text5.Text <> "" Then
Text6.SetFocus
End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text6.Text <> "" Then
Text7.SetFocus
End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text7.Text <> "" Then
Command2.SetFocus
End If
End Sub
