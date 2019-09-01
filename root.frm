VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   10125
   ClientLeft      =   120
   ClientTop       =   1095
   ClientWidth     =   15060
   Icon            =   "root.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10125
   ScaleWidth      =   15060
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Caption         =   "Õ”«»ê—"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2445
      Index           =   1
      Left            =   8535
      TabIndex        =   92
      Top             =   5400
      Width           =   3660
      Begin VB.CommandButton Commandt 
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2715
         TabIndex        =   119
         Top             =   1575
         Width           =   390
      End
      Begin VB.CommandButton Commandz 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2715
         TabIndex        =   118
         Top             =   1215
         Width           =   390
      End
      Begin VB.TextBox Textzt5 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2055
         TabIndex        =   117
         Top             =   1575
         Width           =   570
      End
      Begin VB.TextBox Textzt4 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2055
         TabIndex        =   116
         Top             =   1230
         Width           =   570
      End
      Begin VB.TextBox Textzt3 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1365
         TabIndex        =   114
         Top             =   1395
         Width           =   495
      End
      Begin VB.TextBox Textzt2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   570
         TabIndex        =   112
         Top             =   1605
         Width           =   570
      End
      Begin VB.TextBox Textzt1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   570
         TabIndex        =   111
         Top             =   1245
         Width           =   570
      End
      Begin VB.CommandButton Commandtaf 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2715
         TabIndex        =   110
         Top             =   810
         Width           =   390
      End
      Begin VB.CommandButton Commandgam 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2715
         TabIndex        =   109
         Top             =   450
         Width           =   390
      End
      Begin VB.TextBox Textgam5 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2055
         TabIndex        =   108
         Top             =   450
         Width           =   585
      End
      Begin VB.TextBox Textgam4 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2055
         TabIndex        =   107
         Top             =   810
         Width           =   585
      End
      Begin VB.TextBox Textgam3 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1365
         TabIndex        =   105
         Top             =   630
         Width           =   495
      End
      Begin VB.TextBox Textgam2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   570
         TabIndex        =   103
         Top             =   810
         Width           =   585
      End
      Begin VB.TextBox Textgam 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   570
         TabIndex        =   102
         Top             =   480
         Width           =   585
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         Caption         =   "="
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1875
         TabIndex        =   115
         Top             =   1440
         Width           =   120
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1155
         TabIndex        =   113
         Top             =   1485
         Width           =   90
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         Caption         =   "="
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1875
         TabIndex        =   106
         Top             =   675
         Width           =   120
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1185
         TabIndex        =   104
         Top             =   660
         Width           =   120
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "œÊ „⁄«œ·Â ° œÊ „ÃÂÊ·"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Index           =   0
      Left            =   8535
      TabIndex        =   78
      Top             =   7980
      Width           =   3645
      Begin VB.CommandButton d9 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   2865
         TabIndex        =   91
         Top             =   390
         Width           =   330
      End
      Begin VB.TextBox dc2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2085
         TabIndex        =   86
         Top             =   795
         Width           =   510
      End
      Begin VB.TextBox db2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1155
         TabIndex        =   85
         Top             =   780
         Width           =   660
      End
      Begin VB.TextBox da2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   435
         TabIndex        =   84
         Top             =   765
         Width           =   540
      End
      Begin VB.TextBox dc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2085
         TabIndex        =   83
         Top             =   375
         Width           =   525
      End
      Begin VB.TextBox db 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1155
         TabIndex        =   81
         Top             =   375
         Width           =   675
      End
      Begin VB.TextBox da 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   435
         TabIndex        =   79
         Top             =   375
         Width           =   555
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "y="
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1245
         TabIndex        =   101
         Top             =   1350
         Width           =   210
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "x="
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   285
         TabIndex        =   100
         Top             =   1350
         Width           =   210
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1665
         TabIndex        =   90
         Top             =   1335
         Width           =   45
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   600
         TabIndex        =   89
         Top             =   1335
         Width           =   45
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "y="
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1845
         TabIndex        =   88
         Top             =   855
         Width           =   210
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1005
         TabIndex        =   87
         Top             =   840
         Width           =   90
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "y="
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1845
         TabIndex        =   82
         Top             =   450
         Width           =   210
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1005
         TabIndex        =   80
         Top             =   450
         Width           =   90
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "—”„ »—œ«—"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   12270
      TabIndex        =   64
      Top             =   3480
      Width           =   2745
      Begin VB.CommandButton CommandDC 
         Caption         =   "DC>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1965
         TabIndex        =   76
         Top             =   360
         Width           =   450
      End
      Begin VB.CommandButton CommandDB 
         Caption         =   "DB>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1965
         TabIndex        =   75
         Top             =   1095
         Width           =   465
      End
      Begin VB.CommandButton CommandDA 
         Caption         =   "DA>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1965
         TabIndex        =   74
         Top             =   735
         Width           =   465
      End
      Begin VB.CommandButton CommandCD 
         Caption         =   "CD>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   870
         TabIndex        =   73
         Top             =   720
         Width           =   450
      End
      Begin VB.CommandButton CommandCB 
         Caption         =   "CB>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   870
         TabIndex        =   72
         Top             =   360
         Width           =   450
      End
      Begin VB.CommandButton CommandCA 
         Caption         =   "CA>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   870
         TabIndex        =   71
         Top             =   1095
         Width           =   450
      End
      Begin VB.CommandButton CommandBD 
         Caption         =   "BD>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1395
         TabIndex        =   70
         Top             =   735
         Width           =   450
      End
      Begin VB.CommandButton CommandBC 
         Caption         =   "BC>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1395
         TabIndex        =   69
         Top             =   360
         Width           =   450
      End
      Begin VB.CommandButton CommandBA 
         Caption         =   "BA>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1395
         TabIndex        =   68
         Top             =   1095
         Width           =   450
      End
      Begin VB.CommandButton CommandAD 
         Caption         =   "AD>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   315
         TabIndex        =   67
         Top             =   1095
         Width           =   450
      End
      Begin VB.CommandButton CommandAC 
         Caption         =   "AC>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   315
         TabIndex        =   66
         Top             =   720
         Width           =   450
      End
      Begin VB.CommandButton CommandAB 
         Caption         =   "AB>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   315
         TabIndex        =   65
         Top             =   360
         Width           =   450
      End
   End
   Begin VB.CommandButton Commandvaz1 
      BackColor       =   &H80000007&
      Caption         =   "x,y"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7740
      TabIndex        =   63
      Top             =   7980
      Width           =   435
   End
   Begin VB.CommandButton Commandmoq2 
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7305
      TabIndex        =   62
      Top             =   7980
      Width           =   420
   End
   Begin VB.Frame pro 
      Caption         =   "«„ò«‰« "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   195
      TabIndex        =   48
      Top             =   8340
      Width           =   8220
      Begin VB.Label Label10 
         Caption         =   "—‰ê Å” “„Ì‰Â:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   121
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "⁄«œÌ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3450
         TabIndex        =   99
         Top             =   660
         Width           =   375
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         Caption         =   "ﬁ—„“"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   840
         TabIndex        =   98
         Top             =   660
         Width           =   315
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FF00&
         Caption         =   "”»“"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1185
         TabIndex        =   97
         Top             =   660
         Width           =   300
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF0000&
         Caption         =   "¬»Ì"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1530
         TabIndex        =   96
         Top             =   660
         Width           =   255
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         Caption         =   "“—œ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1860
         TabIndex        =   95
         Top             =   660
         Width           =   195
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "Œ«ò” —Ì"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2130
         TabIndex        =   94
         Top             =   660
         Width           =   690
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackColor       =   &H80000007&
         Caption         =   "„‘òÌ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   2895
         TabIndex        =   93
         Top             =   660
         Width           =   525
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "Œ«ò” —Ì"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6135
         TabIndex        =   55
         Top             =   660
         Width           =   690
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "„‘òÌ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   6870
         TabIndex        =   54
         Top             =   660
         Width           =   525
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         Caption         =   "“—œ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5835
         TabIndex        =   53
         Top             =   645
         Width           =   195
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF0000&
         Caption         =   "¬»Ì"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5475
         TabIndex        =   52
         Top             =   660
         Width           =   255
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FF00&
         Caption         =   "”»“"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5115
         TabIndex        =   51
         Top             =   645
         Width           =   300
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         Caption         =   "ﬁ—„“"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4770
         TabIndex        =   50
         Top             =   660
         Width           =   315
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   ":—‰ê ŒÿÊÿ œ” ê«Â"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6300
         TabIndex        =   49
         Top             =   210
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "„⁄«œ·Â Œÿ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Index           =   1
      Left            =   12270
      TabIndex        =   35
      Top             =   7980
      Width           =   2760
      Begin VB.OptionButton mabdaF 
         Caption         =   "€Ì— „»œ« ê–—"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1140
         TabIndex        =   47
         Top             =   1455
         Width           =   1215
      End
      Begin VB.OptionButton mabdaE 
         Caption         =   "„»œ« ê–—"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   46
         Top             =   1425
         Width           =   945
      End
      Begin VB.OptionButton OpD 
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1500
         TabIndex        =   45
         Top             =   1095
         Width           =   510
      End
      Begin VB.OptionButton Opc 
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1080
         TabIndex        =   44
         Top             =   1095
         Width           =   585
      End
      Begin VB.OptionButton Opb 
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   615
         TabIndex        =   43
         Top             =   1095
         Width           =   525
      End
      Begin VB.OptionButton OpA 
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   42
         Top             =   1095
         Width           =   525
      End
      Begin VB.TextBox Text000 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1335
         TabIndex        =   41
         Top             =   675
         Width           =   540
      End
      Begin VB.TextBox Text5555 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   390
         TabIndex        =   39
         Top             =   690
         Width           =   660
      End
      Begin VB.TextBox Text666 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   390
         TabIndex        =   37
         Top             =   240
         Width           =   645
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "«⁄„«· »—"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2130
         TabIndex        =   56
         Top             =   1080
         Width           =   570
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "x +"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1050
         TabIndex        =   40
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "y="
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   38
         Top             =   750
         Width           =   210
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "x="
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   36
         Top             =   300
         Width           =   210
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "»—œ«— Â«Ì Ê«Õœ „Œ ’« "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2445
      Index           =   0
      Left            =   12270
      TabIndex        =   16
      Top             =   5400
      Width           =   2760
      Begin VB.CommandButton Command11 
         Caption         =   "r-A"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   345
         TabIndex        =   34
         Top             =   2025
         Width           =   540
      End
      Begin VB.CommandButton Command10 
         Caption         =   "r-B"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2040
         TabIndex        =   33
         Top             =   2025
         Width           =   540
      End
      Begin VB.CommandButton Command6 
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1425
         TabIndex        =   32
         Top             =   2025
         Width           =   615
      End
      Begin VB.CommandButton Command5 
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   855
         TabIndex        =   31
         Top             =   2025
         Width           =   555
      End
      Begin VB.TextBox Text1000 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1965
         TabIndex        =   30
         Top             =   1440
         Width           =   600
      End
      Begin VB.TextBox Text900 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1965
         TabIndex        =   29
         Top             =   1110
         Width           =   600
      End
      Begin VB.TextBox Text700 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1980
         TabIndex        =   28
         Top             =   525
         Width           =   570
      End
      Begin VB.TextBox Text600 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1980
         TabIndex        =   27
         Top             =   195
         Width           =   555
      End
      Begin VB.TextBox Text400 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1110
         TabIndex        =   24
         Top             =   1200
         Width           =   585
      End
      Begin VB.TextBox Text300 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1110
         TabIndex        =   23
         Top             =   360
         Width           =   570
      End
      Begin VB.TextBox Text200 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   315
         TabIndex        =   20
         Top             =   1185
         Width           =   540
      End
      Begin VB.TextBox Text100 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   300
         TabIndex        =   19
         Top             =   345
         Width           =   525
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "J ="
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1710
         TabIndex        =   26
         Top             =   1275
         Width           =   240
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "J ="
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1680
         TabIndex        =   25
         Top             =   405
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "i +"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   885
         TabIndex        =   22
         Top             =   1230
         Width           =   195
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "i +"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   885
         TabIndex        =   21
         Top             =   390
         Width           =   195
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "B="
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   45
         TabIndex        =   18
         Top             =   1230
         Width           =   210
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "A="
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   45
         TabIndex        =   17
         Top             =   360
         Width           =   225
      End
   End
   Begin VB.CommandButton Command4 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   14010
      TabIndex        =   15
      Top             =   2865
      Width           =   255
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   13290
      TabIndex        =   13
      Top             =   3090
      Width           =   645
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   13290
      TabIndex        =   12
      Top             =   2730
      Width           =   645
   End
   Begin VB.CommandButton Command3 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   14025
      TabIndex        =   8
      Top             =   1995
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   14040
      TabIndex        =   7
      Top             =   1170
      Width           =   270
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   13305
      TabIndex        =   6
      Top             =   2235
      Width           =   645
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   13305
      TabIndex        =   5
      Top             =   1875
      Width           =   645
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   13320
      TabIndex        =   4
      Top             =   1410
      Width           =   645
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   13320
      TabIndex        =   3
      Top             =   1050
      Width           =   630
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   14055
      TabIndex        =   2
      Top             =   315
      Width           =   255
   End
   Begin VB.TextBox ia 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   13320
      TabIndex        =   1
      Top             =   570
      Width           =   630
   End
   Begin VB.TextBox Ja 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   13320
      TabIndex        =   0
      Top             =   195
      Width           =   630
   End
   Begin VB.PictureBox sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      ScaleHeight     =   210
      ScaleWidth      =   15000
      TabIndex        =   77
      Top             =   9855
      Width           =   15060
   End
   Begin VB.Label Label39 
      AutoSize        =   -1  'True
      Caption         =   "Powerd by NRM team"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   225
      TabIndex        =   120
      Top             =   8100
      Width           =   1545
   End
   Begin VB.Line i 
      X1              =   4440
      X2              =   4440
      Y1              =   465
      Y2              =   6975
   End
   Begin VB.Line j 
      X1              =   7635
      X2              =   1290
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line js 
      BorderColor     =   &H000000FF&
      X1              =   4470
      X2              =   4830
      Y1              =   3705
      Y2              =   3705
   End
   Begin VB.Line is 
      BorderColor     =   &H80000002&
      X1              =   4455
      X2              =   4455
      Y1              =   3690
      Y2              =   3330
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   4770
      X2              =   4770
      Y1              =   3705
      Y2              =   3645
   End
   Begin VB.Line Line2 
      Index           =   0
      X1              =   5130
      X2              =   5130
      Y1              =   3645
      Y2              =   3720
   End
   Begin VB.Line Line3 
      Index           =   0
      X1              =   5490
      X2              =   5490
      Y1              =   3660
      Y2              =   3735
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   5850
      X2              =   5850
      Y1              =   3720
      Y2              =   3660
   End
   Begin VB.Line Line2 
      Index           =   1
      X1              =   6210
      X2              =   6210
      Y1              =   3660
      Y2              =   3735
   End
   Begin VB.Line Line3 
      Index           =   1
      X1              =   6570
      X2              =   6570
      Y1              =   3660
      Y2              =   3735
   End
   Begin VB.Line Line2 
      Index           =   2
      X1              =   6930
      X2              =   6930
      Y1              =   3660
      Y2              =   3735
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   1650
      X2              =   1650
      Y1              =   3750
      Y2              =   3690
   End
   Begin VB.Line Line2 
      Index           =   3
      X1              =   2010
      X2              =   2010
      Y1              =   3690
      Y2              =   3765
   End
   Begin VB.Line Line3 
      Index           =   3
      X1              =   2370
      X2              =   2370
      Y1              =   3690
      Y2              =   3765
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   2730
      X2              =   2730
      Y1              =   3750
      Y2              =   3690
   End
   Begin VB.Line Line2 
      Index           =   4
      X1              =   3090
      X2              =   3090
      Y1              =   3690
      Y2              =   3765
   End
   Begin VB.Line Line3 
      Index           =   4
      X1              =   3450
      X2              =   3450
      Y1              =   3690
      Y2              =   3765
   End
   Begin VB.Line Line2 
      Index           =   5
      X1              =   3810
      X2              =   3810
      Y1              =   3690
      Y2              =   3765
   End
   Begin VB.Line Line3 
      Index           =   5
      X1              =   4170
      X2              =   4170
      Y1              =   3690
      Y2              =   3765
   End
   Begin VB.Line das9 
      X1              =   4410
      X2              =   4485
      Y1              =   3345
      Y2              =   3345
   End
   Begin VB.Line das8 
      X1              =   4410
      X2              =   4485
      Y1              =   2985
      Y2              =   2985
   End
   Begin VB.Line das7 
      X1              =   4410
      X2              =   4485
      Y1              =   2610
      Y2              =   2610
   End
   Begin VB.Line das6 
      X1              =   4410
      X2              =   4485
      Y1              =   2265
      Y2              =   2265
   End
   Begin VB.Line das5 
      X1              =   4410
      X2              =   4485
      Y1              =   1905
      Y2              =   1905
   End
   Begin VB.Line das4 
      X1              =   4410
      X2              =   4485
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line das3 
      X1              =   4410
      X2              =   4485
      Y1              =   1185
      Y2              =   1185
   End
   Begin VB.Line das2 
      X1              =   4410
      X2              =   4485
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line2 
      Index           =   6
      X1              =   7290
      X2              =   7290
      Y1              =   3660
      Y2              =   3735
   End
   Begin VB.Line das18 
      X1              =   4425
      X2              =   4500
      Y1              =   6585
      Y2              =   6585
   End
   Begin VB.Line das17 
      X1              =   4425
      X2              =   4500
      Y1              =   6225
      Y2              =   6225
   End
   Begin VB.Line das16 
      X1              =   4425
      X2              =   4500
      Y1              =   5865
      Y2              =   5865
   End
   Begin VB.Line das15 
      X1              =   4410
      X2              =   4485
      Y1              =   5505
      Y2              =   5505
   End
   Begin VB.Line das14 
      X1              =   4395
      X2              =   4470
      Y1              =   5145
      Y2              =   5145
   End
   Begin VB.Line das13 
      X1              =   4410
      X2              =   4485
      Y1              =   4785
      Y2              =   4785
   End
   Begin VB.Line das11 
      X1              =   4425
      X2              =   4500
      Y1              =   4425
      Y2              =   4425
   End
   Begin VB.Line das10 
      X1              =   4410
      X2              =   4485
      Y1              =   4065
      Y2              =   4065
   End
   Begin VB.Line Line11111 
      Index           =   0
      X1              =   1290
      X2              =   1290
      Y1              =   3675
      Y2              =   3780
   End
   Begin VB.Line Line333333 
      X1              =   7650
      X2              =   7650
      Y1              =   3735
      Y2              =   3645
   End
   Begin VB.Line das1 
      X1              =   4410
      X2              =   4485
      Y1              =   465
      Y2              =   465
   End
   Begin VB.Line das19 
      X1              =   4425
      X2              =   4485
      Y1              =   6945
      Y2              =   6945
   End
   Begin VB.Label d 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   45
      Left            =   4680
      TabIndex        =   61
      Top             =   3735
      Width           =   45
   End
   Begin VB.Label c 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   45
      Left            =   4620
      TabIndex        =   60
      Top             =   3735
      Width           =   45
   End
   Begin VB.Label noqteh 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   45
      Left            =   4545
      TabIndex        =   59
      Top             =   3735
      Width           =   45
   End
   Begin VB.Label a 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   45
      Left            =   4455
      TabIndex        =   58
      Top             =   3735
      Width           =   45
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "x=0 . y=0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5730
      TabIndex        =   57
      Top             =   7995
      Width           =   750
   End
   Begin VB.Label Label000 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0FF&
      Caption         =   "D="
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   13035
      TabIndex        =   14
      Top             =   2940
      Width           =   225
   End
   Begin VB.Label Label000 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "C="
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   13065
      TabIndex        =   11
      Top             =   2085
      Width           =   225
   End
   Begin VB.Label Label000 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "B="
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   13050
      TabIndex        =   10
      Top             =   1260
      Width           =   240
   End
   Begin VB.Label Label000 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "A="
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   13065
      TabIndex        =   9
      Top             =   405
      Width           =   225
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim jm As Integer
Dim im As Integer
Dim jo As Integer
Dim io As Integer
Dim joo As Integer
Dim ioo As Integer
Dim jooo As Integer
Dim iooo As Integer
Dim ra As Integer
Dim rb As Integer
Dim ta As Integer
Dim tb As Integer
Dim X As Integer
Dim Y As Integer
Dim xy As Integer
Dim noqmol As Integer
Dim noqmot As Integer
Dim num1 As Integer
Dim num2 As Integer
Dim nat As Integer
Dim namekar As String

Private Sub Command1_Click()
jm = Ja.Text
im = ia.Text
a.Left = 4440 + jm * 360
a.Top = 3705 - im * 360
End Sub

Private Sub Command10_Click()
Dim w1 As Integer
Dim w2 As Integer
w1 = Text200.Text
w2 = Text400.Text
Line (4440, 3705)-(4440 + w1 * 360, 3705 - w2 * 360)
End Sub

Private Sub Command11_Click()
Dim q1 As Integer
Dim q2 As Integer
q1 = Text100.Text
q2 = Text300.Text
Line (4440, 3705)-(4440 + q1 * 360, 3705 - q2 * 360)
End Sub

Private Sub Command14_Click()
Textsubnum.Text = Textsubnum.Text + "1"
End Sub

Private Sub Command15_Click()
Textsubnum.Text = Textsubnum.Text + "2"
End Sub

Private Sub Command16_Click()
Textsubnum.Text = Textsubnum + "3"
End Sub

Private Sub Command17_Click()
Textsubnum.Text = Textsubnum.Text + "4"
End Sub

Private Sub Command18_Click()
Textsubnum.Text = Textsubnum.Text + "5"
End Sub

Private Sub Command19_Click()
Textsubnum.Text = Textsubnum.Text + "6"
End Sub

Private Sub Command2_Click()
jo = Text3.Text
io = Text4.Text
noqteh.Left = 4440 + jo * 360
noqteh.Top = 3705 - io * 360
End Sub

Private Sub Command20_Click()
Textsubnum.Text = Textsubnum.Text + "7"
End Sub

Private Sub Command21_Click()
Textsubnum.Text = Textsubnum.Text + "8"
End Sub

Private Sub Command22_Click()
Textsubnum.Text = Textsubnum.Text + "9"
End Sub

Private Sub Command23_Click()
Textsubnum.Text = Textsubnum.Text + "0"
End Sub

Private Sub Command24_Click()
num1 = Val(Textsubnum.Text)
op = "+"
Textsubnum.Text = ""
End Sub

Private Sub Command25_Click()
num1 = Textsubnum
num2 = Textsub
nat = num1 + num2
Textsubnum = nat
End Sub

Private Sub Command26_Click()
num1 = Textsubnum
num2 = Textsub
nat = num1 - num2
Textsubnum = nat
End Sub

Private Sub Command27_Click()
num1 = Textsubnum
num2 = Textsub
nat = num1 * num2
Textsubnum = nat
End Sub

Private Sub Command28_Click()
num1 = Textsubnum
num2 = Textsub
nat = num1 / num2
Textsubnum = nat
End Sub

Private Sub Command29_Click()
Textsubnum.Text = ""
num1 = 0
num2 = 0
op = ""
End Sub

Private Sub Command3_Click()
joo = Text5.Text
ioo = Text6.Text
c.Left = 4440 + joo * 360
c.Top = 3705 - ioo * 360
End Sub

Private Sub Command4_Click()
jooo = Text7.Text
iooo = Text8.Text
d.Left = 4440 + jooo * 360
d.Top = 3705 - iooo * 360
End Sub

Private Sub Command5_Click()
ra = Text100.Text
Text600.Text = ra
rb = Text300.Text
Text700.Text = rb
End Sub

Private Sub Command6_Click()
ta = Text200.Text
Text900.Text = ta
tb = Text400.Text
Text1000.Text = tb
End Sub

Private Sub Command999_Click()

End Sub

Private Sub CommandAB_Click()
Line (a.Left, a.Top)-(noqteh.Left, noqteh.Top)
End Sub

Private Sub CommandAC_Click()
Line (a.Left, a.Top)-(c.Left, c.Top)
End Sub

Private Sub CommandAD_Click()
Line (a.Left, a.Top)-(d.Left, d.Top)
End Sub

Private Sub CommandBA_Click()
Line (noqteh.Left, noqteh.Top)-(a.Left, a.Top)
End Sub

Private Sub CommandBC_Click()
Line (noqteh.Left, noqteh.Top)-(c.Left, c.Top)
End Sub

Private Sub CommandBD_Click()
Line (noqteh.Left, noqteh.Top)-(d.Left, d.Top)
End Sub

Private Sub CommandCA_Click()
Line (c.Left, c.Top)-(a.Left, a.Top)
End Sub

Private Sub CommandCB_Click()
Line (c.Left, c.Top)-(noqteh.Left, noqteh.Top)
End Sub

Private Sub CommandCD_Click()
Line (c.Left, c.Top)-(d.Left, d.Top)
End Sub

Private Sub CommandDA_Click()
Line (d.Left, d.Top)-(a.Left, a.Top)
End Sub

Private Sub CommandDB_Click()
Line (d.Left, d.Top)-(noqteh.Left, noqteh.Top)
End Sub

Private Sub CommandDC_Click()
Line (c.Left, c.Top)-(d.Left, d.Top)
End Sub

Private Sub Commandgam_Click()
Dim gam1 As Integer
Dim gam2 As Integer
Dim gam3 As Integer
gam1 = Textgam
gam2 = Textgam3
gam3 = Textgam2
Textgam5.Text = gam1 + gam2
Textgam4.Text = gam3 + gam2
Label26.Caption = "+"
End Sub

Private Sub Commandmoq2_Click()
das1.BorderColor = &H8000000F
das2.BorderColor = &H8000000F
das3.BorderColor = &H8000000F
das4.BorderColor = &H8000000F
das5.BorderColor = &H8000000F
das6.BorderColor = &H8000000F
das7.BorderColor = &H8000000F
das8.BorderColor = &H8000000F
das9.BorderColor = &H8000000F
das10.BorderColor = &H8000000F
das11.BorderColor = &H8000000F
das13.BorderColor = &H8000000F
das14.BorderColor = &H8000000F
das15.BorderColor = &H8000000F
das16.BorderColor = &H8000000F
das17.BorderColor = &H8000000F
das18.BorderColor = &H8000000F
das19.BorderColor = &H8000000F
i.BorderColor = &H8000000F
ia.Text = "1"
ia.Enabled = False
Text4.Text = "1"
Text4.Enabled = False
Text6.Text = "1"
Text6.Enabled = False
Text8.Text = "1"
Text8.Enabled = False
End Sub

Private Sub Commandt_Click()
Textzt4.Text = Textzt1.Text / Textzt3.Text
Textzt5.Text = Textzt2.Text / Textzt3.Text
Label37.Caption = "/"
End Sub

Private Sub Commandtaf_Click()
Textgam5.Text = Textgam.Text - Textgam3.Text
Textgam4.Text = Textgam2.Text - Textgam3.Text
Label26.Caption = "-"
End Sub

Private Sub Commandvaz1_Click()
das1.BorderColor = &H80000007
das2.BorderColor = &H80000007
das3.BorderColor = &H80000007
das4.BorderColor = &H80000007
das5.BorderColor = &H80000007
das6.BorderColor = &H80000007
das7.BorderColor = &H80000007
das8.BorderColor = &H80000007
das9.BorderColor = &H80000007
das10.BorderColor = &H80000007
das11.BorderColor = &H80000007
das13.BorderColor = &H80000007
das14.BorderColor = &H80000007
das15.BorderColor = &H80000007
das16.BorderColor = &H80000007
das17.BorderColor = &H80000007
das18.BorderColor = &H80000007
das19.BorderColor = &H80000007
i.BorderColor = &H80000007
ia.Text = ""
ia.Enabled = True
Text4.Text = ""
Text4.Enabled = True
Text6.Text = ""
Text6.Enabled = True
Text8.Text = ""
Text8.Enabled = True
End Sub

Private Sub Commandz_Click()
Textzt4.Text = Textzt1.Text * Textzt3.Text
Textzt5.Text = Textzt2.Text * Textzt3.Text
Label37.Caption = "*"
End Sub

Private Sub d9_Click()
Dim gavy As Integer
Dim gavx As Integer
gavy = (dc2 - da2 * dc / da) / (db2 - da2 * db / da)
Label25.Caption = gavy
gavx = (dc - db * gavy) / da
Label24 = gavx
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label20.Caption = "x=" + Str(X) + ",y=" + Str(Y)
End Sub

Private Sub Label13_Click()
i.BorderColor = &HFF&
j.BorderColor = &HFF&
End Sub

Private Sub Label14_Click()
i.BorderColor = &HFF00&
j.BorderColor = &HFF00&
End Sub

Private Sub Label15_Click()
i.BorderColor = &HFF0000
j.BorderColor = &HFF0000
End Sub

Private Sub Label16_Click()
i.BorderColor = &HFFFF&
j.BorderColor = &HFFFF&
End Sub

Private Sub Label18_Click()
i.BorderColor = &H80000009
j.BorderColor = &H80000009
End Sub

Private Sub Label19_Click()
i.BorderColor = &H8000000A
j.BorderColor = &H8000000A
End Sub

Private Sub Label27_Click()
Me.BackColor = &H80000007
End Sub

Private Sub Label28_Click()
Me.BackColor = &H8000000A
End Sub

Private Sub Label29_Click()
Me.BackColor = &HFFFF&
End Sub

Private Sub Label30_Click()
Me.BackColor = &HFF0000
End Sub

Private Sub Label31_Click()
Me.BackColor = &HFF00&
End Sub

Private Sub Label32_Click()
Me.BackColor = &HFF&
End Sub

Private Sub Label33_Click()
Me.BackColor = &H8000000F
End Sub

Private Sub mabdaE_Click()
Text000.Text = 0
Text000.Enabled = False
End Sub

Private Sub mabdaF_Click()
Text000.Text = ""
Text000.Enabled = True
End Sub

Private Sub OpA_Click()
Dim X As Integer
Dim Y As Integer
Dim arz As Integer
Dim z As Integer
If Text000.Text = "" Then
Text000.Text = "0"
End If
If Text5555.Text = "" Then
Text5555.Text = "0"
End If
If Text666.Text = "" Then
Text666.Text = "0"
End If
X = Text666.Text
z = Text5555.Text
arz = Text000.Text
Y = z * X + arz
noqmol = 4440 + X * 360
noqmot = 3705 - Y * 360
a.Left = noqmol
a.Top = noqmot
End Sub

Private Sub Opb_Click()
Dim X As Integer
Dim Y As Integer
Dim arz As Integer
Dim z As Integer
If Text000.Text = "" Then
Text000.Text = "0"
End If
If Text5555.Text = "" Then
Text5555.Text = "0"
End If
If Text666.Text = "" Then
Text666.Text = "0"
End If
X = Text666.Text
z = Text5555.Text
arz = Text000.Text
Y = z * X + arz
noqmol = 4440 + X * 360
noqmot = 3705 - Y * 360
noqteh.Left = noqmol
noqteh.Top = noqmot
End Sub

Private Sub Opc_Click()
Dim X As Integer
Dim Y As Integer
Dim arz As Integer
Dim z As Integer
If Text000.Text = "" Then
Text000.Text = "0"
End If
If Text5555.Text = "" Then
Text5555.Text = "0"
End If
If Text666.Text = "" Then
Text666.Text = "0"
End If
X = Text666.Text
z = Text5555.Text
arz = Text000.Text
Y = z * X + arz
noqmol = 4440 + X * 360
noqmot = 3705 - Y * 360
c.Left = noqmol
c.Top = noqmot
End Sub

Private Sub OpD_Click()
Dim X As Integer
Dim Y As Integer
Dim arz As Integer
Dim z As Integer
If Text000.Text = "" Then
Text000.Text = "0"
End If
If Text5555.Text = "" Then
Text5555.Text = "0"
End If
If Text666.Text = "" Then
Text666.Text = "0"
End If
X = Text666.Text
z = Text5555.Text
arz = Text000.Text
Y = z * X + arz
noqmol = 4440 + X * 360
noqmot = 3705 - Y * 360
d.Left = noqmol
d.Top = noqmot
End Sub

Private Sub raby_Click()
i1.BorderColor = &HFF0000
j1.BorderColor = &HFF0000
End Sub

Private Sub rmash_Click()
i1.BorderColor = &H80000007
j1.BorderColor = &H80000007
End Sub

Private Sub rsabz_Click()
i1.BorderColor = &HFF00&
j1.BorderColor = &HFF00&
End Sub

Private Sub rsefid_Click()
i1.BorderColor = &HFFFFFF
j1.BorderColor = &HFFFFFF
End Sub

Private Sub pro_Change()

End Sub

Private Sub TabStrip1_Change()

End Sub
