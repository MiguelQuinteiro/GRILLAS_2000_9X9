VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form txtFilaRegionGrilla 
   BackColor       =   &H00FFC0C0&
   Caption         =   "GRILLAS 2000"
   ClientHeight    =   6885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11565
   LinkTopic       =   "Form1"
   ScaleHeight     =   6885
   ScaleWidth      =   11565
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Slider Slider1 
      Height          =   555
      Left            =   240
      TabIndex        =   150
      Top             =   6000
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   979
      _Version        =   393216
      Max             =   46656
   End
   Begin VB.TextBox txtSalto 
      Alignment       =   1  'Right Justify
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
      Left            =   10800
      TabIndex        =   148
      Text            =   "1"
      Top             =   5280
      Width           =   495
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar1 
      Height          =   495
      Left            =   6720
      TabIndex        =   147
      Top             =   5280
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   873
      _Version        =   393216
      Enabled         =   0   'False
      Arrows          =   65536
      Orientation     =   1179649
   End
   Begin VB.TextBox txtAzul 
      BackColor       =   &H00FFFF80&
      Height          =   495
      Left            =   10800
      TabIndex        =   146
      Top             =   -120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   81
      Left            =   10800
      TabIndex        =   145
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   80
      Left            =   10320
      TabIndex        =   144
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   79
      Left            =   9840
      TabIndex        =   143
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   78
      Left            =   9240
      TabIndex        =   142
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   77
      Left            =   8760
      TabIndex        =   141
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   76
      Left            =   8280
      TabIndex        =   140
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   75
      Left            =   7680
      TabIndex        =   139
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   74
      Left            =   7200
      TabIndex        =   138
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   73
      Left            =   6720
      TabIndex        =   137
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   72
      Left            =   10800
      TabIndex        =   136
      Top             =   4080
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   71
      Left            =   10320
      TabIndex        =   135
      Top             =   4080
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   70
      Left            =   9840
      TabIndex        =   134
      Top             =   4080
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   69
      Left            =   9240
      TabIndex        =   133
      Top             =   4080
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   68
      Left            =   8760
      TabIndex        =   132
      Top             =   4080
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   67
      Left            =   8280
      TabIndex        =   131
      Top             =   4080
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   66
      Left            =   7680
      TabIndex        =   130
      Top             =   4080
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   65
      Left            =   7200
      TabIndex        =   129
      Top             =   4080
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   64
      Left            =   6720
      TabIndex        =   128
      Top             =   4080
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   63
      Left            =   10800
      TabIndex        =   127
      Top             =   3600
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   62
      Left            =   10320
      TabIndex        =   126
      Top             =   3600
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   61
      Left            =   9840
      TabIndex        =   125
      Top             =   3600
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   60
      Left            =   9240
      TabIndex        =   124
      Top             =   3600
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   59
      Left            =   8760
      TabIndex        =   123
      Top             =   3600
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   58
      Left            =   8280
      TabIndex        =   122
      Top             =   3600
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   57
      Left            =   7680
      TabIndex        =   121
      Top             =   3600
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   56
      Left            =   7200
      TabIndex        =   120
      Top             =   3600
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   55
      Left            =   6720
      TabIndex        =   119
      Top             =   3600
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   54
      Left            =   10800
      TabIndex        =   118
      Top             =   3000
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   53
      Left            =   10320
      TabIndex        =   117
      Top             =   3000
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   52
      Left            =   9840
      TabIndex        =   116
      Top             =   3000
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   51
      Left            =   9240
      TabIndex        =   115
      Top             =   3000
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   50
      Left            =   8760
      TabIndex        =   114
      Top             =   3000
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   49
      Left            =   8280
      TabIndex        =   113
      Top             =   3000
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   48
      Left            =   7680
      TabIndex        =   112
      Top             =   3000
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   47
      Left            =   7200
      TabIndex        =   111
      Top             =   3000
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   46
      Left            =   6720
      TabIndex        =   110
      Top             =   3000
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   45
      Left            =   10800
      TabIndex        =   109
      Top             =   2520
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   44
      Left            =   10320
      TabIndex        =   108
      Top             =   2520
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   43
      Left            =   9840
      TabIndex        =   107
      Top             =   2520
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   42
      Left            =   9240
      TabIndex        =   106
      Top             =   2520
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   41
      Left            =   8760
      TabIndex        =   105
      Top             =   2520
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   40
      Left            =   8280
      TabIndex        =   104
      Top             =   2520
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   39
      Left            =   7680
      TabIndex        =   103
      Top             =   2520
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   38
      Left            =   7200
      TabIndex        =   102
      Top             =   2520
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   37
      Left            =   6720
      TabIndex        =   101
      Top             =   2520
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   36
      Left            =   10800
      TabIndex        =   100
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   35
      Left            =   10320
      TabIndex        =   99
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   34
      Left            =   9840
      TabIndex        =   98
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   33
      Left            =   9240
      TabIndex        =   97
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   32
      Left            =   8760
      TabIndex        =   96
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   31
      Left            =   8280
      TabIndex        =   95
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   30
      Left            =   7680
      TabIndex        =   94
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   29
      Left            =   7200
      TabIndex        =   93
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   28
      Left            =   6720
      TabIndex        =   92
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   27
      Left            =   10800
      TabIndex        =   91
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   26
      Left            =   10320
      TabIndex        =   90
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   25
      Left            =   9840
      TabIndex        =   89
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   24
      Left            =   9240
      TabIndex        =   88
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   23
      Left            =   8760
      TabIndex        =   87
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   22
      Left            =   8280
      TabIndex        =   86
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   21
      Left            =   7680
      TabIndex        =   85
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   20
      Left            =   7200
      TabIndex        =   84
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   19
      Left            =   6720
      TabIndex        =   83
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   18
      Left            =   10800
      TabIndex        =   82
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   17
      Left            =   10320
      TabIndex        =   81
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   16
      Left            =   9840
      TabIndex        =   80
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   15
      Left            =   9240
      TabIndex        =   79
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   14
      Left            =   8760
      TabIndex        =   78
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   13
      Left            =   8280
      TabIndex        =   77
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   12
      Left            =   7680
      TabIndex        =   76
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   11
      Left            =   7200
      TabIndex        =   75
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   10
      Left            =   6720
      TabIndex        =   74
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   10800
      TabIndex        =   73
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   10320
      TabIndex        =   72
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   9840
      TabIndex        =   71
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   9240
      TabIndex        =   70
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   8760
      TabIndex        =   69
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   8280
      TabIndex        =   68
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   7680
      TabIndex        =   67
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   7200
      TabIndex        =   66
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   6720
      TabIndex        =   65
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtBlanco 
      Height          =   495
      Left            =   10320
      TabIndex        =   64
      Top             =   -120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtColumnaRegionGrilla 
      Alignment       =   2  'Center
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
      Index           =   9
      Left            =   5760
      TabIndex        =   62
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox txtColumnaRegionGrilla 
      Alignment       =   2  'Center
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
      Index           =   8
      Left            =   5280
      TabIndex        =   61
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox txtColumnaRegionGrilla 
      Alignment       =   2  'Center
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
      Index           =   7
      Left            =   4800
      TabIndex        =   60
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox txtColumnaRegionGrilla 
      Alignment       =   2  'Center
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
      Index           =   6
      Left            =   4320
      TabIndex        =   59
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox txtColumnaRegionGrilla 
      Alignment       =   2  'Center
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
      Index           =   5
      Left            =   3840
      TabIndex        =   58
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox txtColumnaRegionGrilla 
      Alignment       =   2  'Center
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
      Index           =   4
      Left            =   3360
      TabIndex        =   57
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox txtColumnaRegionGrilla 
      Alignment       =   2  'Center
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
      Index           =   3
      Left            =   2880
      TabIndex        =   56
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox txtColumnaRegionGrilla 
      Alignment       =   2  'Center
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
      Index           =   2
      Left            =   2400
      TabIndex        =   55
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox txtColumnaRegionGrilla 
      Alignment       =   2  'Center
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
      Index           =   1
      Left            =   1920
      TabIndex        =   54
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox txtFilaRegionGrilla 
      Alignment       =   2  'Center
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
      Index           =   9
      Left            =   5760
      TabIndex        =   53
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txtFilaRegionGrilla 
      Alignment       =   2  'Center
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
      Index           =   8
      Left            =   5280
      TabIndex        =   52
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txtFilaRegionGrilla 
      Alignment       =   2  'Center
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
      Index           =   7
      Left            =   4800
      TabIndex        =   51
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txtFilaRegionGrilla 
      Alignment       =   2  'Center
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
      Index           =   6
      Left            =   4320
      TabIndex        =   50
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txtFilaRegionGrilla 
      Alignment       =   2  'Center
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
      Index           =   5
      Left            =   3840
      TabIndex        =   49
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txtFilaRegionGrilla 
      Alignment       =   2  'Center
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
      Index           =   4
      Left            =   3360
      TabIndex        =   48
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txtFilaRegionGrilla 
      Alignment       =   2  'Center
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
      Index           =   3
      Left            =   2880
      TabIndex        =   47
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txtFilaRegionGrilla 
      Alignment       =   2  'Center
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
      Index           =   2
      Left            =   2400
      TabIndex        =   46
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txtFilaRegionGrilla 
      Alignment       =   2  'Center
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
      Index           =   1
      Left            =   1920
      TabIndex        =   45
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txtRegionGrilla 
      Alignment       =   2  'Center
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
      Index           =   9
      Left            =   5760
      TabIndex        =   42
      Top             =   3120
      Width           =   495
   End
   Begin VB.TextBox txtRegionGrilla 
      Alignment       =   2  'Center
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
      Index           =   8
      Left            =   5280
      TabIndex        =   41
      Top             =   3120
      Width           =   495
   End
   Begin VB.TextBox txtRegionGrilla 
      Alignment       =   2  'Center
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
      Index           =   7
      Left            =   4800
      TabIndex        =   40
      Top             =   3120
      Width           =   495
   End
   Begin VB.TextBox txtRegionGrilla 
      Alignment       =   2  'Center
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
      Index           =   6
      Left            =   4320
      TabIndex        =   39
      Top             =   3120
      Width           =   495
   End
   Begin VB.TextBox txtRegionGrilla 
      Alignment       =   2  'Center
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
      Index           =   5
      Left            =   3840
      TabIndex        =   38
      Top             =   3120
      Width           =   495
   End
   Begin VB.TextBox txtRegionGrilla 
      Alignment       =   2  'Center
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
      Index           =   4
      Left            =   3360
      TabIndex        =   37
      Top             =   3120
      Width           =   495
   End
   Begin VB.TextBox txtRegionGrilla 
      Alignment       =   2  'Center
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
      Index           =   3
      Left            =   2880
      TabIndex        =   36
      Top             =   3120
      Width           =   495
   End
   Begin VB.TextBox txtRegionGrilla 
      Alignment       =   2  'Center
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
      Index           =   2
      Left            =   2400
      TabIndex        =   35
      Top             =   3120
      Width           =   495
   End
   Begin VB.TextBox txtRegionGrilla 
      Alignment       =   2  'Center
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
      Index           =   1
      Left            =   1920
      TabIndex        =   34
      Top             =   3120
      Width           =   495
   End
   Begin VB.TextBox txtColumnaGrilla 
      Alignment       =   2  'Center
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
      Index           =   9
      Left            =   5760
      TabIndex        =   32
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox txtColumnaGrilla 
      Alignment       =   2  'Center
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
      Index           =   8
      Left            =   5280
      TabIndex        =   31
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox txtColumnaGrilla 
      Alignment       =   2  'Center
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
      Index           =   7
      Left            =   4800
      TabIndex        =   30
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox txtColumnaGrilla 
      Alignment       =   2  'Center
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
      Index           =   6
      Left            =   4320
      TabIndex        =   29
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox txtColumnaGrilla 
      Alignment       =   2  'Center
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
      Index           =   5
      Left            =   3840
      TabIndex        =   28
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox txtColumnaGrilla 
      Alignment       =   2  'Center
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
      Index           =   4
      Left            =   3360
      TabIndex        =   27
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox txtColumnaGrilla 
      Alignment       =   2  'Center
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
      Index           =   3
      Left            =   2880
      TabIndex        =   26
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox txtColumnaGrilla 
      Alignment       =   2  'Center
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
      Index           =   2
      Left            =   2400
      TabIndex        =   25
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox txtColumnaGrilla 
      Alignment       =   2  'Center
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
      Index           =   1
      Left            =   1920
      TabIndex        =   24
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox txtFilaGrilla 
      Alignment       =   2  'Center
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
      Index           =   9
      Left            =   5760
      TabIndex        =   22
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox txtFilaGrilla 
      Alignment       =   2  'Center
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
      Index           =   8
      Left            =   5280
      TabIndex        =   21
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox txtFilaGrilla 
      Alignment       =   2  'Center
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
      Index           =   7
      Left            =   4800
      TabIndex        =   20
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox txtFilaGrilla 
      Alignment       =   2  'Center
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
      Index           =   6
      Left            =   4320
      TabIndex        =   19
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox txtFilaGrilla 
      Alignment       =   2  'Center
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
      Index           =   5
      Left            =   3840
      TabIndex        =   18
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox txtFilaGrilla 
      Alignment       =   2  'Center
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
      Index           =   4
      Left            =   3360
      TabIndex        =   17
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox txtFilaGrilla 
      Alignment       =   2  'Center
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
      Index           =   3
      Left            =   2880
      TabIndex        =   16
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox txtFilaGrilla 
      Alignment       =   2  'Center
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
      Index           =   2
      Left            =   2400
      TabIndex        =   15
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox txtDigitoGrilla 
      Alignment       =   2  'Center
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
      Index           =   9
      Left            =   5760
      TabIndex        =   14
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtFilaGrilla 
      Alignment       =   2  'Center
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
      Index           =   1
      Left            =   1920
      TabIndex        =   13
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox txtDigitoGrilla 
      Alignment       =   2  'Center
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
      Index           =   8
      Left            =   5280
      TabIndex        =   10
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtDigitoGrilla 
      Alignment       =   2  'Center
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
      Index           =   7
      Left            =   4800
      TabIndex        =   9
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtDigitoGrilla 
      Alignment       =   2  'Center
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
      Index           =   6
      Left            =   4320
      TabIndex        =   8
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtDigitoGrilla 
      Alignment       =   2  'Center
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
      Index           =   5
      Left            =   3840
      TabIndex        =   7
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtDigitoGrilla 
      Alignment       =   2  'Center
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
      Index           =   4
      Left            =   3360
      TabIndex        =   6
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtDigitoGrilla 
      Alignment       =   2  'Center
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
      Index           =   3
      Left            =   2880
      TabIndex        =   5
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtDigitoGrilla 
      Alignment       =   2  'Center
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
      Index           =   2
      Left            =   2400
      TabIndex        =   4
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtDigitoGrilla 
      Alignment       =   2  'Center
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
      Index           =   1
      Left            =   1920
      TabIndex        =   3
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtNumeroGrilla 
      Alignment       =   1  'Right Justify
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
      Left            =   1920
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdCargaGrillas 
      BackColor       =   &H00FF8080&
      Caption         =   "CARGA LAS GRILLAS EN MEMORIA"
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5280
      Width           =   6015
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFC0C0&
      Caption         =   "SALTO"
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
      Left            =   9960
      TabIndex        =   149
      Top             =   5280
      Width           =   735
   End
   Begin VB.Line Line1 
      X1              =   6480
      X2              =   6480
      Y1              =   120
      Y2              =   6000
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Coloca un número entre 1 y 46656"
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
      Left            =   3360
      TabIndex        =   63
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Caption         =   "Columna en Región del Dígito"
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
      Left            =   240
      TabIndex        =   44
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Caption         =   "Fila en Región del Dígito"
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
      Left            =   240
      TabIndex        =   43
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Caption         =   "Región del Dígito"
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
      Left            =   240
      TabIndex        =   33
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Caption         =   "Columna del Dígito"
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
      Left            =   240
      TabIndex        =   23
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Caption         =   "Fila del Dígito"
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
      Left            =   240
      TabIndex        =   12
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Caption         =   "Digitos de Grilla"
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
      Left            =   240
      TabIndex        =   11
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Caption         =   "Número de Grilla"
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
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "txtFilaRegionGrilla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************************
'* PROYECTO      : GRILLAS 2000
'* CONTENIDO     : PERMITE AVISUALIZAR LAS GRILLAS EN UNA SOLUCION DE SUDOKU
'* VERSION       : 1.1
'* AUTORES       : MIGUEL QUINTEIRO PIÑERO / MIGUEL QUINTEIRO FERNANDEZ
'* INICIO        : 04 DE MARZO DE 2014
'* ACTUALIZACION : 04 DE MARZO DE 2014
'****************************************************************************************
Option Explicit

Option Base 1

Private Type LasGrillas
  NumeroGrilla As Long
  DigitoGrilla(1 To 9) As Integer
  FilaGrilla(1 To 9) As Integer
  ColumnaGrilla(1 To 9) As Integer
  RegionGrilla(1 To 9) As Integer
  FilaRegionGrilla(1 To 9) As Integer
  ColumnaRegionGrilla(1 To 9) As Integer
End Type

Dim miGrilla(1 To 46656) As LasGrillas

Dim miLineInput As String
Dim miLineOutput As String
Dim miNumeroGrilla As Long

Private Sub cmdCargaGrillas_Click()
  Dim i As Integer

  Dim miDigito1 As Integer
  Dim miDigito2 As Integer
  Dim miDigito3 As Integer
  Dim miDigito4 As Integer
  Dim miDigito5 As Integer
  Dim miDigito6 As Integer
  Dim miDigito7 As Integer
  Dim miDigito8 As Integer
  Dim miDigito9 As Integer

  Open "LasGrillas.txt" For Input As #10
  Do Until EOF(10)
    Line Input #10, miLineInput
    miNumeroGrilla = Val(Mid(miLineInput, 35, 5))

    miGrilla(miNumeroGrilla).NumeroGrilla = Val(Mid(miLineInput, 35, 5))

    miGrilla(miNumeroGrilla).DigitoGrilla(1) = Val(Mid(miLineInput, 2, 2))
    miGrilla(miNumeroGrilla).DigitoGrilla(2) = Val(Mid(miLineInput, 5, 2))
    miGrilla(miNumeroGrilla).DigitoGrilla(3) = Val(Mid(miLineInput, 8, 2))
    miGrilla(miNumeroGrilla).DigitoGrilla(4) = Val(Mid(miLineInput, 11, 2))
    miGrilla(miNumeroGrilla).DigitoGrilla(5) = Val(Mid(miLineInput, 14, 2))
    miGrilla(miNumeroGrilla).DigitoGrilla(6) = Val(Mid(miLineInput, 17, 2))
    miGrilla(miNumeroGrilla).DigitoGrilla(7) = Val(Mid(miLineInput, 20, 2))
    miGrilla(miNumeroGrilla).DigitoGrilla(8) = Val(Mid(miLineInput, 23, 2))
    miGrilla(miNumeroGrilla).DigitoGrilla(9) = Val(Mid(miLineInput, 26, 2))

    For i = 1 To 9
      miGrilla(miNumeroGrilla).FilaGrilla(i) = (Int(((miGrilla(miNumeroGrilla).DigitoGrilla(i)) - 1) / 9) + 1)
      miGrilla(miNumeroGrilla).ColumnaGrilla(i) = ((((miGrilla(miNumeroGrilla).DigitoGrilla(i)) - 1) Mod 9) + 1)
      ' Región 1
      If miGrilla(miNumeroGrilla).FilaGrilla(i) >= 1 And miGrilla(miNumeroGrilla).FilaGrilla(i) <= 3 And _
         miGrilla(miNumeroGrilla).ColumnaGrilla(i) >= 1 And miGrilla(miNumeroGrilla).ColumnaGrilla(i) <= 3 Then
        miGrilla(miNumeroGrilla).RegionGrilla(i) = 1
      End If
      ' Región 2
      If miGrilla(miNumeroGrilla).FilaGrilla(i) >= 1 And miGrilla(miNumeroGrilla).FilaGrilla(i) <= 3 And _
         miGrilla(miNumeroGrilla).ColumnaGrilla(i) >= 4 And miGrilla(miNumeroGrilla).ColumnaGrilla(i) <= 6 Then
        miGrilla(miNumeroGrilla).RegionGrilla(i) = 2
      End If
      ' Región 3
      If miGrilla(miNumeroGrilla).FilaGrilla(i) >= 1 And miGrilla(miNumeroGrilla).FilaGrilla(i) <= 3 And _
         miGrilla(miNumeroGrilla).ColumnaGrilla(i) >= 7 And miGrilla(miNumeroGrilla).ColumnaGrilla(i) <= 9 Then
        miGrilla(miNumeroGrilla).RegionGrilla(i) = 3
      End If
      ' Región 4
      If miGrilla(miNumeroGrilla).FilaGrilla(i) >= 4 And miGrilla(miNumeroGrilla).FilaGrilla(i) <= 6 And _
         miGrilla(miNumeroGrilla).ColumnaGrilla(i) >= 1 And miGrilla(miNumeroGrilla).ColumnaGrilla(i) <= 3 Then
        miGrilla(miNumeroGrilla).RegionGrilla(i) = 4
      End If
      ' Región 5
      If miGrilla(miNumeroGrilla).FilaGrilla(i) >= 4 And miGrilla(miNumeroGrilla).FilaGrilla(i) <= 6 And _
         miGrilla(miNumeroGrilla).ColumnaGrilla(i) >= 4 And miGrilla(miNumeroGrilla).ColumnaGrilla(i) <= 6 Then
        miGrilla(miNumeroGrilla).RegionGrilla(i) = 5
      End If
      ' Región 6
      If miGrilla(miNumeroGrilla).FilaGrilla(i) >= 4 And miGrilla(miNumeroGrilla).FilaGrilla(i) <= 6 And _
         miGrilla(miNumeroGrilla).ColumnaGrilla(i) >= 7 And miGrilla(miNumeroGrilla).ColumnaGrilla(i) <= 9 Then
        miGrilla(miNumeroGrilla).RegionGrilla(i) = 6
      End If
      ' Región 7
      If miGrilla(miNumeroGrilla).FilaGrilla(i) >= 7 And miGrilla(miNumeroGrilla).FilaGrilla(i) <= 9 And _
         miGrilla(miNumeroGrilla).ColumnaGrilla(i) >= 1 And miGrilla(miNumeroGrilla).ColumnaGrilla(i) <= 3 Then
        miGrilla(miNumeroGrilla).RegionGrilla(i) = 7
      End If
      ' Región 8
      If miGrilla(miNumeroGrilla).FilaGrilla(i) >= 7 And miGrilla(miNumeroGrilla).FilaGrilla(i) <= 9 And _
         miGrilla(miNumeroGrilla).ColumnaGrilla(i) >= 4 And miGrilla(miNumeroGrilla).ColumnaGrilla(i) <= 6 Then
        miGrilla(miNumeroGrilla).RegionGrilla(i) = 8
      End If
      ' Región 9
      If miGrilla(miNumeroGrilla).FilaGrilla(i) >= 7 And miGrilla(miNumeroGrilla).FilaGrilla(i) <= 9 And _
         miGrilla(miNumeroGrilla).ColumnaGrilla(i) >= 7 And miGrilla(miNumeroGrilla).ColumnaGrilla(i) <= 9 Then
        miGrilla(miNumeroGrilla).RegionGrilla(i) = 9
      End If

      If ((Int((miGrilla(miNumeroGrilla).DigitoGrilla(i) - 1) / 9) + 1) Mod 3) = 1 Then
        miGrilla(miNumeroGrilla).FilaRegionGrilla(i) = 1
      End If
      If ((Int((miGrilla(miNumeroGrilla).DigitoGrilla(i) - 1) / 9) + 1) Mod 3) = 2 Then
        miGrilla(miNumeroGrilla).FilaRegionGrilla(i) = 2
      End If
      If ((Int((miGrilla(miNumeroGrilla).DigitoGrilla(i) - 1) / 9) + 1) Mod 3) = 0 Then
        miGrilla(miNumeroGrilla).FilaRegionGrilla(i) = 3
      End If

      If ((miGrilla(miNumeroGrilla).DigitoGrilla(i)) Mod 3) = 1 Then
        miGrilla(miNumeroGrilla).ColumnaRegionGrilla(i) = 1
      End If
      If ((miGrilla(miNumeroGrilla).DigitoGrilla(i)) Mod 3) = 2 Then
        miGrilla(miNumeroGrilla).ColumnaRegionGrilla(i) = 2
      End If
      If ((miGrilla(miNumeroGrilla).DigitoGrilla(i)) Mod 3) = 0 Then
        miGrilla(miNumeroGrilla).ColumnaRegionGrilla(i) = 3
      End If
    Next i
  Loop
  Close #10
  cmdCargaGrillas.Enabled = False
  FlatScrollBar1.Enabled = True
End Sub

Private Sub FlatScrollBar1_Change()
  txtNumeroGrilla = Val(txtSalto) * FlatScrollBar1.Value
End Sub

Private Sub Slider1_Click()
  txtNumeroGrilla = Val(txtSalto) * Slider1.Value
End Sub

Private Sub txtNumeroGrilla_Change()
  Dim i As Integer
  Dim x As Integer

  If txtNumeroGrilla <> "" And Val(txtNumeroGrilla) > 0 And Val(txtNumeroGrilla) < 46657 Then
    miNumeroGrilla = Val(txtNumeroGrilla)
    For x = 1 To 81
      txtMuestraGrilla(x).BackColor = txtBlanco.BackColor
      txtMuestraGrilla(x).Text = ""
    Next x
    For i = 1 To 9
      With miGrilla(miNumeroGrilla)
        txtDigitoGrilla(i) = .DigitoGrilla(i)
        txtFilaGrilla(i) = .FilaGrilla(i)
        txtColumnaGrilla(i) = .ColumnaGrilla(i)
        txtRegionGrilla(i) = .RegionGrilla(i)
        txtFilaRegionGrilla(i) = .FilaRegionGrilla(i)
        txtColumnaRegionGrilla(i) = .ColumnaRegionGrilla(i)

        txtMuestraGrilla(.DigitoGrilla(i)).BackColor = txtAzul.BackColor
        txtMuestraGrilla(.DigitoGrilla(i)).Text = .DigitoGrilla(i)
      End With
    Next i
  Else
    For i = 1 To 9
      txtDigitoGrilla(i) = ""
      txtFilaGrilla(i) = ""
      txtColumnaGrilla(i) = ""
      txtRegionGrilla(i) = ""
      txtFilaRegionGrilla(i) = ""
      txtColumnaRegionGrilla(i) = ""
    Next i
    For x = 1 To 81
      txtMuestraGrilla(x).BackColor = txtBlanco.BackColor
      txtMuestraGrilla(x).Text = ""
    Next x
  End If
End Sub
