VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form7 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Form7"
   ClientHeight    =   8595
   ClientLeft      =   4755
   ClientTop       =   1050
   ClientWidth     =   13080
   FillStyle       =   3  'Vertical Line
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "login_Form7.frx":0000
   ScaleHeight     =   8595
   ScaleWidth      =   13080
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Caption         =   "Log in"
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
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7080
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      DataSource      =   "Adodc1"
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
      IMEMode         =   3  'DISABLE
      Left            =   6120
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   6480
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      DataSource      =   "Adodc1"
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
      Left            =   6120
      TabIndex        =   0
      Top             =   6000
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   -7560
      Top             =   0
      Visible         =   0   'False
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   1296
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
   Begin VB.Label Label3 
      BackColor       =   &H8000000D&
      Caption         =   "<-- EXIT"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   8160
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   1335
      Left            =   5880
      Picture         =   "login_Form7.frx":C92A
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000002&
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
      Height          =   375
      Left            =   4800
      TabIndex        =   4
      Top             =   6480
      Width           =   1455
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
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
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   8580
      Left            =   0
      Picture         =   "login_Form7.frx":1417F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13080
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim conn As New ADODB.Connection

Private Sub Command1_Click()
Form8.Visible = True
Form7.Visible = False
End Sub

Private Sub Command2_Click()
If rs.State <> adStateClosed Then
rs.Close
End If
If Text1.Text = "" Then
MsgBox "Enter the playerID", vbInformation + vbOKOnly, "login"
Text1.SetFocus
Exit Sub
End If

If Text2.Text = "" Then
MsgBox "Enter the password", vbInformation + vbOKOnly, "login"
Text2.SetFocus
Exit Sub
End If

If Text1.Text <> "" And Text2.Text <> "" Then
If rs.State = 1 Then
rs.Close
Else
rs.Open "select * from tableLogin where playerID='" & Text1 & "' and Password='" & Text2 & "'", conn, adOpenStatic, adLockOptimistic, adCmdText
End If

If rs.EOF = True Then
MsgBox "Invalid playerID or password", vbCritical + vbOKOnly, "login"
Text1.Text = ""
Text2.Text = ""
Text1.SetFocus
Else
MsgBox "Login successful", vbInformation + vbOKOnly, "login"
Form1.Visible = True
Unload Me
End If
End If
End Sub

Private Sub Form_Load()
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Ruby Kumari\Pictures\LOGIN-DATABSE.MDB;Persist Security Info=False"
'conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\home\Documents\vb project\LOGIN-DATABSE.MDB;Persist Security Info=False"
End Sub

Private Sub Label3_Click()
End
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text1.Text <> "" Then
Text2.SetFocus
End If
End Sub


Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text2.Text <> "" Then
Command2.SetFocus
End If
End Sub


