VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form16 
   Caption         =   "Form16"
   ClientHeight    =   7065
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12675
   LinkTopic       =   "Form16"
   ScaleHeight     =   7065
   ScaleWidth      =   12675
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   480
      Top             =   7200
      Width           =   5655
      _ExtentX        =   9975
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
   Begin VB.Frame Frame3 
      Height          =   6975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12615
      Begin VB.Frame Frame4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5415
         Left            =   2400
         TabIndex        =   1
         Top             =   840
         Width           =   7935
         Begin VB.CommandButton Command1 
            Caption         =   "UPDATE"
            Enabled         =   0   'False
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
            Left            =   3240
            TabIndex        =   13
            Top             =   4680
            Width           =   1695
         End
         Begin VB.TextBox Text16 
            DataSource      =   "Adodc1"
            Height          =   495
            Left            =   3720
            TabIndex        =   8
            Top             =   720
            Width           =   2535
         End
         Begin VB.TextBox Text11 
            DataSource      =   "Adodc1"
            Enabled         =   0   'False
            Height          =   525
            IMEMode         =   3  'DISABLE
            Left            =   3120
            PasswordChar    =   "*"
            TabIndex        =   7
            Top             =   3840
            Width           =   2535
         End
         Begin VB.TextBox Text10 
            DataSource      =   "Adodc1"
            Enabled         =   0   'False
            Height          =   495
            IMEMode         =   3  'DISABLE
            Left            =   3120
            PasswordChar    =   "*"
            TabIndex        =   6
            Top             =   3240
            Width           =   2535
         End
         Begin VB.TextBox Text9 
            DataSource      =   "Adodc1"
            Enabled         =   0   'False
            Height          =   495
            Left            =   3120
            TabIndex        =   5
            Top             =   2640
            Width           =   2535
         End
         Begin VB.CommandButton Command6 
            Caption         =   "BACK"
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
            Left            =   720
            TabIndex        =   4
            Top             =   4680
            Width           =   1815
         End
         Begin VB.CommandButton Command5 
            Caption         =   "SEARCH"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   3360
            TabIndex        =   3
            Top             =   1440
            Width           =   1695
         End
         Begin VB.CommandButton Command4 
            Caption         =   "DELETE"
            Enabled         =   0   'False
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
            Left            =   5520
            TabIndex        =   2
            Top             =   4680
            Width           =   1695
         End
         Begin VB.Label Label5 
            Caption         =   "Search"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   17.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3360
            TabIndex        =   14
            Top             =   0
            Width           =   1335
         End
         Begin VB.Label Label23 
            Caption         =   "Enter the player name:"
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
            Left            =   1680
            TabIndex        =   12
            Top             =   840
            Width           =   2175
         End
         Begin VB.Line Line6 
            X1              =   0
            X2              =   7935
            Y1              =   2400
            Y2              =   2415
         End
         Begin VB.Label Label18 
            Caption         =   "Name:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2280
            TabIndex        =   11
            Top             =   2760
            Width           =   615
         End
         Begin VB.Label Label17 
            Caption         =   "playerID:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2160
            TabIndex        =   10
            Top             =   3360
            Width           =   855
         End
         Begin VB.Label Label16 
            Caption         =   "Password:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2040
            TabIndex        =   9
            Top             =   3960
            Width           =   975
         End
         Begin VB.Line Line4 
            X1              =   0
            X2              =   7935
            Y1              =   4560
            Y2              =   4575
         End
      End
      Begin VB.Image Image2 
         Height          =   6855
         Left            =   0
         Top             =   120
         Width           =   12615
      End
   End
End
Attribute VB_Name = "Form16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim conn As New ADODB.Connection

Private Sub Command1_Click()
Load Form14
Form14.Show
rs.Close
conn.Close
Unload Me
End Sub

Private Sub Command4_Click()
Load Form15
Form15.Show
rs.Close
conn.Close
Unload Me
End Sub

Private Sub Command5_Click()
If rs.State <> adStateClosed Then
rs.Close
End If
If Text16.Text = "" Then
MsgBox "Enter the playerID", vbInformation + vbOKOnly, "Search"
Text16.SetFocus
Exit Sub
End If

If Text16.Text <> "" Then
If rs.State = 1 Then
rs.Close
Else
rs.Open "select * from tableLogin where Name='" & Text16 & "'", conn, adOpenStatic, adLockOptimistic, adCmdText
End If

If rs.EOF = True Then
MsgBox "Player not found", vbCritical + vbOKOnly, "Search"
Text16.Text = ""
Text16.SetFocus
Else
Command4.Enabled = True
Command1.Enabled = True
Text9.Text = rs(0)
Text10.Text = rs(1)
Text11.Text = rs(2)
End If
End If
End Sub

Private Sub Command6_Click()
Load Form1
Form1.Show
rs.Close
conn.Close
Unload Me
End Sub


Private Sub Form_Load()
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Ruby Kumari\Pictures\LOGIN-DATABSE.MDB;Persist Security Info=False"
'conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\home\Documents\vb project\LOGIN-DATABSE.MDB;Persist Security Info=False"
rs.Open "select * from tableLogin", conn, adOpenDynamic, adLockOptimistic, adCmdText
End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text16.Text <> "" Then
Command5.SetFocus
ElseIf (KeyAscii < 65 And KeyAscii <> 8 And KeyAscii > 32) Or (KeyAscii > 90 And KeyAscii < 97) Or (KeyAscii > 122) Then
KeyAscii = 0
MsgBox "Enter characters only"
End If
End Sub
