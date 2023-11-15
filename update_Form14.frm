VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form14 
   Caption         =   "Form14"
   ClientHeight    =   7305
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12885
   LinkTopic       =   "Form14"
   ScaleHeight     =   7305
   ScaleWidth      =   12885
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   12615
      Begin VB.Frame Frame2 
         Height          =   5415
         Left            =   4560
         TabIndex        =   3
         Top             =   840
         Width           =   7935
         Begin VB.CommandButton Command3 
            Caption         =   "SUBMIT"
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
            TabIndex        =   25
            Top             =   4680
            Width           =   1695
         End
         Begin VB.CommandButton Command1 
            Caption         =   "VERIFY"
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
            Left            =   5640
            TabIndex        =   24
            Top             =   720
            Width           =   1695
         End
         Begin VB.CommandButton Command2 
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
            Left            =   960
            TabIndex        =   23
            Top             =   4680
            Width           =   1815
         End
         Begin VB.TextBox Text8 
            DataSource      =   "Adodc1"
            Height          =   495
            Left            =   5160
            TabIndex        =   19
            Top             =   3240
            Width           =   2535
         End
         Begin VB.TextBox Text7 
            DataSource      =   "Adodc1"
            Height          =   495
            Left            =   5160
            TabIndex        =   18
            Top             =   3840
            Width           =   2535
         End
         Begin VB.TextBox Text6 
            DataSource      =   "Adodc1"
            Height          =   495
            Left            =   5160
            TabIndex        =   15
            Top             =   2640
            Width           =   2535
         End
         Begin VB.TextBox Text5 
            Enabled         =   0   'False
            Height          =   495
            Left            =   1200
            TabIndex        =   13
            Top             =   3840
            Width           =   2535
         End
         Begin VB.TextBox Text4 
            Enabled         =   0   'False
            Height          =   495
            Left            =   1200
            TabIndex        =   11
            Top             =   3240
            Width           =   2535
         End
         Begin VB.TextBox Text3 
            Enabled         =   0   'False
            Height          =   495
            Left            =   1200
            TabIndex        =   9
            Top             =   2640
            Width           =   2535
         End
         Begin VB.TextBox Text2 
            DataSource      =   "Adodc1"
            Height          =   495
            Left            =   2400
            TabIndex        =   7
            Top             =   1200
            Width           =   2535
         End
         Begin VB.TextBox Text1 
            DataSource      =   "Adodc1"
            Height          =   495
            Left            =   2400
            TabIndex        =   5
            Top             =   480
            Width           =   2535
         End
         Begin VB.Label Label2 
            Caption         =   "Update"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   17.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3240
            TabIndex        =   22
            Top             =   -120
            Width           =   1335
         End
         Begin VB.Label Label13 
            Caption         =   "NEW"
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
            Left            =   5760
            TabIndex        =   21
            Top             =   2280
            Width           =   495
         End
         Begin VB.Label Label12 
            Caption         =   "OLD"
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
            Left            =   1920
            TabIndex        =   20
            Top             =   2280
            Width           =   375
         End
         Begin VB.Line Line3 
            X1              =   0
            X2              =   7935
            Y1              =   4560
            Y2              =   4575
         End
         Begin VB.Label Label11 
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
            Left            =   4200
            TabIndex        =   17
            Top             =   3960
            Width           =   975
         End
         Begin VB.Label Label10 
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
            Left            =   4200
            TabIndex        =   16
            Top             =   3360
            Width           =   855
         End
         Begin VB.Label Label9 
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
            Left            =   4200
            TabIndex        =   14
            Top             =   2760
            Width           =   615
         End
         Begin VB.Line Line2 
            X1              =   3960
            X2              =   3975
            Y1              =   2400
            Y2              =   4575
         End
         Begin VB.Line Line1 
            X1              =   0
            X2              =   7935
            Y1              =   2400
            Y2              =   2415
         End
         Begin VB.Label Label8 
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
            Left            =   240
            TabIndex        =   12
            Top             =   3960
            Width           =   975
         End
         Begin VB.Label Label7 
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
            Left            =   240
            TabIndex        =   10
            Top             =   3360
            Width           =   855
         End
         Begin VB.Label Label6 
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
            Left            =   240
            TabIndex        =   8
            Top             =   2760
            Width           =   615
         End
         Begin VB.Label Label4 
            Caption         =   "Enter the password of above Player"
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
            Left            =   120
            TabIndex        =   6
            Top             =   1200
            Width           =   2175
         End
         Begin VB.Label Label3 
            Caption         =   "Enter the playerID of the player u want to upadate"
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
            Left            =   120
            TabIndex        =   4
            Top             =   480
            Width           =   2175
         End
      End
      Begin VB.PictureBox DataGrid1 
         DragMode        =   1  'Automatic
         Height          =   3495
         Left            =   1200
         ScaleHeight     =   3435
         ScaleWidth      =   1995
         TabIndex        =   1
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "List signed-up Players"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   5295
      End
      Begin VB.Image Image1 
         Height          =   6855
         Left            =   0
         Top             =   120
         Width           =   12615
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   -720
      Top             =   6960
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim conn As New ADODB.Connection

Private Sub Command1_Click()
If rs.State <> adStateClosed Then
rs.Close
End If
If Text1.Text = "" Then
MsgBox "Enter the playerID", vbInformation + vbOKOnly, "update"
Text1.SetFocus
Exit Sub
End If

If Text2.Text = "" Then
MsgBox "Enter the password", vbInformation + vbOKOnly, "update"
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
MsgBox "Invalid playerID or password", vbCritical + vbOKOnly, "update"
Text1.Text = ""
Text2.Text = ""
Text1.SetFocus
Else
MsgBox "Verifiation successfull", vbInformation + vbOKOnly, "update"
Command3.Enabled = True
Text3.Text = rs(0)
Text4.Text = rs(1)
Text5.Text = rs(2)
End If
End If
End Sub

Private Sub Command2_Click()
Load Form1
Form1.Show
rs.Close
conn.Close
Unload Me
End Sub

Private Sub Command3_Click()
Dim update As Boolean
If Text6.Text <> "" Then
rs("Name") = Text6.Text
update = True
End If
If Text8.Text <> "" Then
rs("playerID") = Text8.Text
update = True
End If
If Text7.Text <> "" Then
rs("Password") = Text7.Text
update = True
End If
rs.update
Adodc1.Refresh
If update = True Then
MsgBox "Updated succesfully", vbInformation + vbOKOnly, "update"
Else
MsgBox "Update failed", vbCritical + vbOKOnly, "update"
End If
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Command3.Enabled = False
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text1.Text <> "" Then
Text2.SetFocus
End If
End Sub

Private Sub Form_Load()
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Ruby Kumari\Pictures\LOGIN-DATABSE.MDB;Persist Security Info=False"
'conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\home\Documents\vb project\LOGIN-DATABSE.MDB;Persist Security Info=False"
rs.Open "select * from tableLogin", conn, adOpenDynamic, adLockOptimistic, adCmdText
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text2.Text <> "" Then
Command1.SetFocus
End If
End Sub
Private Sub Text6_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13 And Text6.Text <> "") Or (KeyAscii = 13 And Text6.Text = "") Then
Text8.SetFocus
ElseIf (KeyAscii < 65 And KeyAscii <> 8 And KeyAscii > 32) Or (KeyAscii > 90 And KeyAscii < 97) Or (KeyAscii > 122) Then
KeyAscii = 0
MsgBox "Enter characters only"
End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13 And Text7.Text <> "") Or (KeyAscii = 13 And Text7.Text = "") Then
Command3.SetFocus
End If
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13 And Text8.Text <> "") Or (KeyAscii = 13 And Text8.Text = "") Then
Text7.SetFocus
End If
End Sub
