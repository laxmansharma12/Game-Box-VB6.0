VERSION 5.00
Begin VB.Form Form9 
   Caption         =   "Form9"
   ClientHeight    =   10935
   ClientLeft      =   3210
   ClientTop       =   1815
   ClientWidth     =   18210
   LinkTopic       =   "Form9"
   ScaleHeight     =   10935
   ScaleWidth      =   18210
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "FINISH"
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
      Left            =   6000
      TabIndex        =   2
      Top             =   6720
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7695
      Left            =   11640
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      Begin VB.Timer Timer1 
         Interval        =   900
         Left            =   3720
         Top             =   1080
      End
      Begin VB.Label Label127 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   30
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   1680
         TabIndex        =   1
         Top             =   2880
         Width           =   1095
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   0
      TabIndex        =   128
      Top             =   120
      Width           =   5415
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C000C0&
      Height          =   7335
      Left            =   0
      TabIndex        =   127
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C000C0&
      Height          =   7575
      Left            =   11280
      TabIndex        =   126
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   125
      Top             =   7440
      Width           =   11175
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   5880
      TabIndex        =   124
      Top             =   120
      Width           =   5415
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   240
      TabIndex        =   123
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C000C0&
      Height          =   1575
      Left            =   1920
      TabIndex        =   122
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C000C0&
      Height          =   2055
      Left            =   4560
      TabIndex        =   121
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   4680
      TabIndex        =   120
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   5640
      TabIndex        =   119
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C000C0&
      Height          =   855
      Left            =   6360
      TabIndex        =   118
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C000C0&
      Height          =   495
      Left            =   5640
      TabIndex        =   117
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   4560
      TabIndex        =   116
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C000C0&
      Height          =   855
      Left            =   5040
      TabIndex        =   115
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label15 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   4800
      TabIndex        =   114
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C000C0&
      Height          =   855
      Left            =   5400
      TabIndex        =   113
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label17 
      BackColor       =   &H00C000C0&
      Height          =   855
      Left            =   6240
      TabIndex        =   112
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label Label18 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   5400
      TabIndex        =   111
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label20 
      BackColor       =   &H00C000C0&
      Height          =   2055
      Left            =   7440
      TabIndex        =   110
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label21 
      BackColor       =   &H00C000C0&
      Height          =   1935
      Left            =   8160
      TabIndex        =   109
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label22 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   7440
      TabIndex        =   108
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label23 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   6720
      TabIndex        =   107
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label24 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   8280
      TabIndex        =   106
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label25 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   9480
      TabIndex        =   105
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C000C0&
      Height          =   1935
      Left            =   10680
      TabIndex        =   104
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label27 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   10800
      TabIndex        =   103
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label28 
      BackColor       =   &H00C000C0&
      Height          =   1095
      Left            =   9360
      TabIndex        =   102
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Label19 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   6240
      TabIndex        =   101
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label29 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   6480
      TabIndex        =   100
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label30 
      BackColor       =   &H00C000C0&
      Height          =   1935
      Left            =   10200
      TabIndex        =   99
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label31 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   9120
      TabIndex        =   98
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label32 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   9480
      TabIndex        =   97
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label33 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   1200
      TabIndex        =   96
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label34 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   1440
      TabIndex        =   95
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label35 
      BackColor       =   &H00C000C0&
      Height          =   615
      Left            =   2400
      TabIndex        =   94
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Label36 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   2400
      TabIndex        =   93
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label37 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   0
      TabIndex        =   92
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label38 
      BackColor       =   &H00C000C0&
      Height          =   855
      Left            =   600
      TabIndex        =   91
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label39 
      BackColor       =   &H00C000C0&
      Height          =   855
      Left            =   3840
      TabIndex        =   90
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label40 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   2760
      TabIndex        =   89
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label41 
      BackColor       =   &H00C000C0&
      Height          =   855
      Left            =   3240
      TabIndex        =   88
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label42 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   3480
      TabIndex        =   87
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label43 
      BackColor       =   &H00C000C0&
      Height          =   855
      Left            =   3960
      TabIndex        =   86
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label44 
      BackColor       =   &H00C000C0&
      Height          =   855
      Left            =   2520
      TabIndex        =   85
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label45 
      BackColor       =   &H00C000C0&
      Height          =   2295
      Left            =   1320
      TabIndex        =   84
      Top             =   5160
      Width           =   255
   End
   Begin VB.Label Label46 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   720
      TabIndex        =   83
      Top             =   6960
      Width           =   615
   End
   Begin VB.Label Label47 
      BackColor       =   &H00C000C0&
      Height          =   615
      Left            =   720
      TabIndex        =   82
      Top             =   6360
      Width           =   255
   End
   Begin VB.Label Label48 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   480
      TabIndex        =   81
      Top             =   6120
      Width           =   495
   End
   Begin VB.Label Label49 
      BackColor       =   &H00C000C0&
      Height          =   2295
      Left            =   1800
      TabIndex        =   80
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label50 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   1080
      TabIndex        =   79
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label Label51 
      BackColor       =   &H00C000C0&
      Height          =   2295
      Left            =   2520
      TabIndex        =   78
      Top             =   3840
      Width           =   255
   End
   Begin VB.Label Label52 
      BackColor       =   &H00C000C0&
      Height          =   2295
      Left            =   3120
      TabIndex        =   77
      Top             =   5400
      Width           =   255
   End
   Begin VB.Label Label53 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   3120
      TabIndex        =   76
      Top             =   5160
      Width           =   735
   End
   Begin VB.Label Label54 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   2520
      TabIndex        =   75
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label Label55 
      BackColor       =   &H00C000C0&
      Height          =   855
      Left            =   840
      TabIndex        =   74
      Top             =   4320
      Width           =   255
   End
   Begin VB.Label Label56 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   480
      TabIndex        =   73
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label Label57 
      BackColor       =   &H00C000C0&
      Height          =   1575
      Left            =   480
      TabIndex        =   72
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label58 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   480
      TabIndex        =   71
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label59 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   1080
      TabIndex        =   70
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label Label60 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   1800
      TabIndex        =   69
      Top             =   4920
      Width           =   735
   End
   Begin VB.Label Label61 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   2040
      TabIndex        =   68
      Top             =   5520
      Width           =   495
   End
   Begin VB.Label Label62 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   1440
      TabIndex        =   67
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Label Label63 
      BackColor       =   &H00C000C0&
      Height          =   615
      Left            =   2160
      TabIndex        =   66
      Top             =   6600
      Width           =   255
   End
   Begin VB.Label Label64 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   1920
      TabIndex        =   65
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label65 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   3360
      TabIndex        =   64
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label66 
      BackColor       =   &H00C000C0&
      Height          =   735
      Left            =   4560
      TabIndex        =   63
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label67 
      BackColor       =   &H00C000C0&
      Height          =   495
      Left            =   3360
      TabIndex        =   62
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label68 
      BackColor       =   &H00C000C0&
      Height          =   855
      Left            =   8760
      TabIndex        =   61
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label69 
      BackColor       =   &H00C000C0&
      Height          =   495
      Left            =   3720
      TabIndex        =   60
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label70 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   8880
      TabIndex        =   59
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label71 
      BackColor       =   &H00C000C0&
      Height          =   1335
      Left            =   4080
      TabIndex        =   58
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label72 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   4920
      TabIndex        =   57
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label Label73 
      BackColor       =   &H00C000C0&
      Height          =   855
      Left            =   4920
      TabIndex        =   56
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Label74 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   4920
      TabIndex        =   55
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label Label75 
      BackColor       =   &H00C000C0&
      Height          =   855
      Left            =   3600
      TabIndex        =   54
      Top             =   5760
      Width           =   255
   End
   Begin VB.Label Label76 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   3600
      TabIndex        =   53
      Top             =   5760
      Width           =   735
   End
   Begin VB.Label Label77 
      BackColor       =   &H00C000C0&
      Height          =   855
      Left            =   4080
      TabIndex        =   52
      Top             =   5760
      Width           =   255
   End
   Begin VB.Label Label78 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   4080
      TabIndex        =   51
      Top             =   6480
      Width           =   1695
   End
   Begin VB.Label Label79 
      BackColor       =   &H00C000C0&
      Height          =   1335
      Left            =   5520
      TabIndex        =   50
      Top             =   6120
      Width           =   255
   End
   Begin VB.Label Label80 
      BackColor       =   &H00C000C0&
      Height          =   615
      Left            =   3360
      TabIndex        =   49
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label Label81 
      BackColor       =   &H00C000C0&
      Height          =   615
      Left            =   2400
      TabIndex        =   48
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label82 
      BackColor       =   &H00C000C0&
      Height          =   615
      Left            =   4680
      TabIndex        =   47
      Top             =   6960
      Width           =   255
   End
   Begin VB.Label Label83 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   3240
      TabIndex        =   46
      Top             =   6960
      Width           =   735
   End
   Begin VB.Label Label84 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   3960
      TabIndex        =   45
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label Label85 
      BackColor       =   &H00C000C0&
      Height          =   855
      Left            =   6000
      TabIndex        =   44
      Top             =   5400
      Width           =   255
   End
   Begin VB.Label Label86 
      BackColor       =   &H00C000C0&
      Height          =   1575
      Left            =   6000
      TabIndex        =   43
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Label88 
      BackColor       =   &H00C000C0&
      Height          =   615
      Left            =   5160
      TabIndex        =   42
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Label89 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   5760
      TabIndex        =   41
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label90 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   6480
      TabIndex        =   40
      Top             =   6000
      Width           =   2535
   End
   Begin VB.Label Label91 
      BackColor       =   &H00C000C0&
      Height          =   1935
      Left            =   8040
      TabIndex        =   39
      Top             =   5520
      Width           =   255
   End
   Begin VB.Label Label92 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   5520
      TabIndex        =   38
      Top             =   6000
      Width           =   495
   End
   Begin VB.Label Label93 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   5880
      TabIndex        =   37
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label Label94 
      BackColor       =   &H00C000C0&
      Height          =   615
      Left            =   8640
      TabIndex        =   36
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label95 
      BackColor       =   &H00C000C0&
      Height          =   495
      Left            =   4680
      TabIndex        =   35
      Top             =   6000
      Width           =   255
   End
   Begin VB.Label Label96 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   8520
      TabIndex        =   34
      Top             =   6720
      Width           =   735
   End
   Begin VB.Label Label97 
      BackColor       =   &H00C000C0&
      Height          =   615
      Left            =   9000
      TabIndex        =   33
      Top             =   6840
      Width           =   255
   End
   Begin VB.Label Label98 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   8760
      TabIndex        =   32
      Top             =   6360
      Width           =   735
   End
   Begin VB.Label Label99 
      BackColor       =   &H00C000C0&
      Height          =   1215
      Left            =   9240
      TabIndex        =   31
      Top             =   5280
      Width           =   255
   End
   Begin VB.Label Label100 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   7680
      TabIndex        =   30
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label Label101 
      BackColor       =   &H00C000C0&
      Height          =   855
      Left            =   8520
      TabIndex        =   29
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Label102 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   8760
      TabIndex        =   28
      Top             =   5280
      Width           =   495
   End
   Begin VB.Label Label103 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   5880
      TabIndex        =   27
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label104 
      BackColor       =   &H00C000C0&
      Height          =   855
      Left            =   6840
      TabIndex        =   26
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Label105 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   6720
      TabIndex        =   25
      Top             =   5520
      Width           =   735
   End
   Begin VB.Label Label106 
      BackColor       =   &H00C000C0&
      Height          =   1575
      Left            =   6840
      TabIndex        =   24
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label107 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   7920
      TabIndex        =   23
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label108 
      BackColor       =   &H00C000C0&
      Height          =   855
      Left            =   7920
      TabIndex        =   22
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label Label109 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   7440
      TabIndex        =   21
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Label110 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   6960
      TabIndex        =   20
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label Label111 
      BackColor       =   &H00C000C0&
      Height          =   615
      Left            =   7440
      TabIndex        =   19
      Top             =   3840
      Width           =   255
   End
   Begin VB.Label Label112 
      BackColor       =   &H00C000C0&
      Height          =   1095
      Left            =   8760
      TabIndex        =   18
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label113 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   8280
      TabIndex        =   17
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label114 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   9480
      TabIndex        =   16
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label115 
      BackColor       =   &H00C000C0&
      Height          =   855
      Left            =   9960
      TabIndex        =   15
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Label116 
      BackColor       =   &H00C000C0&
      Height          =   1095
      Left            =   10440
      TabIndex        =   14
      Top             =   6000
      Width           =   255
   End
   Begin VB.Label Label117 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   9840
      TabIndex        =   13
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Label Label118 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   10440
      TabIndex        =   12
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Label Label119 
      BackColor       =   &H00C000C0&
      Height          =   615
      Left            =   9840
      TabIndex        =   11
      Top             =   6840
      Width           =   255
   End
   Begin VB.Label Label120 
      BackColor       =   &H00C000C0&
      Height          =   495
      Left            =   9240
      TabIndex        =   10
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label Label121 
      BackColor       =   &H00C000C0&
      Height          =   1215
      Left            =   10800
      TabIndex        =   9
      Top             =   4560
      Width           =   255
   End
   Begin VB.Label Label122 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   9720
      TabIndex        =   8
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label123 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   9480
      TabIndex        =   7
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label Label124 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   10440
      TabIndex        =   6
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label125 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   10320
      TabIndex        =   5
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label Label126 
      BackColor       =   &H00C000C0&
      Height          =   735
      Left            =   9720
      TabIndex        =   4
      Top             =   5280
      Width           =   255
   End
   Begin VB.Label Label87 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   5280
      TabIndex        =   3
      Top             =   4080
      Width           =   735
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MsgBox "You Win"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub

Private Sub Label127_Click()
Timer1.Enabled = True
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label14_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label15_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label16_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label17_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label18_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label19_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label20_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label21_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label22_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label23_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label24_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label25_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label26_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label27_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label28_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label29_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label30_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label31_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label32_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label33_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label34_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label35_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label36_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label37_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label38_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label39_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label40_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label41_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label42_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label43_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label44_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label45_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label46_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label47_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label48_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label49_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label50_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label51_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label52_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label53_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label54_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label55_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label56_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label57_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label58_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label59_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label60_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label61_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label62_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label63_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label64_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label65_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label66_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label67_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label68_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label69_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label70_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label71_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label72_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label73_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label74_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label75_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label76_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label77_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label78_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label79_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Labe80_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label81_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label82_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label83_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label84_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label85_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label86_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label87_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label88_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label89_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label90_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label91_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label92_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label93_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label94_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label95_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label96_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label97_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label98_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label99_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label100_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label101_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label102_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label103_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label104_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label105_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label106_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label107_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label108_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label109_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Labe110_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label111_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label112_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label113_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label114_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label115_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label116_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label117_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label118_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label119_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label120_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label121_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label122_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label123_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label124_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label125_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub
Private Sub Label126_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Game Over"
Form10.Visible = True
Unload Me
End Sub

Private Sub Timer1_Timer()
Label127.Caption = Label127.Caption + 1
End Sub

