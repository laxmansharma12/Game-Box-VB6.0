VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form17 
   Caption         =   "Form17"
   ClientHeight    =   7245
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13635
   LinkTopic       =   "Form17"
   ScaleHeight     =   7245
   ScaleWidth      =   13635
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   4935
      Left            =   2280
      TabIndex        =   16
      Top             =   960
      Width           =   7815
      Begin VB.CommandButton Command5 
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
         Left            =   3000
         TabIndex        =   19
         Top             =   3720
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   3600
         TabIndex        =   18
         Top             =   2400
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   3600
         TabIndex        =   17
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Image Image1 
         Height          =   4815
         Left            =   0
         Top             =   120
         Width           =   7815
      End
      Begin VB.Label Label2 
         Caption         =   "This is only for Administrator "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         TabIndex        =   22
         Top             =   240
         Width           =   5295
      End
      Begin VB.Label Label4 
         Caption         =   "Enter the password of Admin:"
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
         Left            =   960
         TabIndex        =   21
         Top             =   2520
         Width           =   2535
      End
      Begin VB.Label Label3 
         Caption         =   "Enter Admin ID:"
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
         Left            =   2160
         TabIndex        =   20
         Top             =   1560
         Width           =   2175
      End
   End
   Begin VB.Frame Frame3 
      Height          =   6975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   12615
      Begin VB.Frame Frame4 
         Height          =   5055
         Left            =   5520
         TabIndex        =   3
         Top             =   840
         Width           =   6615
         Begin VB.CommandButton Command3 
            Caption         =   "DOWN"
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
            Left            =   2400
            TabIndex        =   15
            Top             =   4200
            Width           =   1815
         End
         Begin VB.CommandButton Command2 
            Caption         =   "UPDATE"
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
            Left            =   480
            TabIndex        =   14
            Top             =   3360
            Width           =   1695
         End
         Begin VB.CommandButton Command1 
            Caption         =   "UP"
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
            Left            =   2400
            TabIndex        =   13
            Top             =   2520
            Width           =   1815
         End
         Begin VB.TextBox Text11 
            DataField       =   "Password"
            DataSource      =   "Adodc1"
            Height          =   525
            Left            =   2520
            TabIndex        =   8
            Top             =   1560
            Width           =   2535
         End
         Begin VB.TextBox Text10 
            DataField       =   "playerID"
            DataSource      =   "Adodc1"
            Height          =   495
            Left            =   2520
            TabIndex        =   7
            Top             =   960
            Width           =   2535
         End
         Begin VB.TextBox Text9 
            DataField       =   "Name"
            DataSource      =   "Adodc1"
            Height          =   495
            Left            =   2520
            TabIndex        =   6
            Top             =   360
            Width           =   2535
         End
         Begin VB.CommandButton Command6 
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
            Height          =   615
            Left            =   2400
            TabIndex        =   5
            Top             =   3360
            Width           =   1815
         End
         Begin VB.CommandButton Command4 
            Caption         =   "DELETE"
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
            Left            =   4440
            TabIndex        =   4
            Top             =   3360
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "DETAILS"
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
            Left            =   2880
            TabIndex        =   12
            Top             =   0
            Width           =   1215
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
            Left            =   1800
            TabIndex        =   11
            Top             =   480
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
            Left            =   1560
            TabIndex        =   10
            Top             =   1080
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
            Left            =   1560
            TabIndex        =   9
            Top             =   1680
            Width           =   975
         End
         Begin VB.Line Line4 
            X1              =   0
            X2              =   7935
            Y1              =   2280
            Y2              =   2295
         End
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   495
         Left            =   8400
         Top             =   8400
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   873
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
      Begin VB.PictureBox DataGrid2 
         DragMode        =   1  'Automatic
         Height          =   5895
         Left            =   120
         ScaleHeight     =   5835
         ScaleWidth      =   4995
         TabIndex        =   1
         Top             =   960
         Width           =   5055
      End
      Begin VB.Image Image2 
         Height          =   6855
         Left            =   0
         Top             =   120
         Width           =   12615
      End
      Begin VB.Label Label24 
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
   End
End
Attribute VB_Name = "Form17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim conn As New ADODB.Connection

Private Sub Command1_Click()
Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF Then
Adodc1.Recordset.MoveFirst
MsgBox "Reached First Record"
End If
End Sub

Private Sub Command2_Click()
rs(0) = Text9.Text
rs(1) = Text10.Text
rs(2) = Text11.Text
rs.update
Adodc1.Refresh
MsgBox "Updated Successfully"
End Sub

Private Sub Command3_Click()
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF Then
Adodc1.Recordset.MoveLast
MsgBox "Reached Last Record"
End If
End Sub

Private Sub Command4_Click()
Dim wish As Integer
wish = MsgBox("Do you want to delete (Y/N)?", vbQuestion + vbYesNo + vbDefaultButton1)
If wish = vbYes Then
rs.Delete
rs.MoveNext
MsgBox "Account Deleted succesfully", vbInformation + vbOKOnly, "Delete"
Else
MsgBox "Account Deletion failed", vbCritical + vbOKOnly, "Delete"
End If
Adodc1.Refresh
End Sub

Private Sub Command5_Click()
If rs.State <> adStateClosed Then
rs.Close
End If
If Text1.Text = "" Then
MsgBox "Enter the Admin ID", vbInformation + vbOKOnly, "update"
Text1.SetFocus
Exit Sub
End If

If Text2.Text = "" Then
MsgBox "Enter the Admin password", vbInformation + vbOKOnly, "update"
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
MsgBox "Invalid Admin ID or password", vbCritical + vbOKOnly, "update"
Text1.Text = ""
Text2.Text = ""
Text1.SetFocus
Else
MsgBox "Verifiation successfull", vbInformation + vbOKOnly, "update"
Frame3.Visible = True
Frame1.Visible = False
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

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text1.Text <> "" Then
Text2.SetFocus
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text2.Text <> "" Then
Command5.SetFocus
End If
End Sub
