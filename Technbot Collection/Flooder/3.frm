VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form scroll 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Flooding"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "3.frx":0000
   ScaleHeight     =   5955
   ScaleWidth      =   8475
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox cust9 
      Height          =   285
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   74
      Top             =   5280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox cust3 
      Height          =   285
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   73
      Top             =   5280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox cust4 
      Height          =   285
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   72
      Top             =   5280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox cust5 
      Height          =   285
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   71
      Top             =   5280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox cust6 
      Height          =   285
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   70
      Top             =   5280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox cust7 
      Height          =   285
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   69
      Top             =   5280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox cust8 
      Height          =   285
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   68
      Top             =   5280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox cust2 
      Height          =   285
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   67
      Top             =   5280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox cust1 
      Height          =   285
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   66
      Top             =   5280
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComDlg.CommonDialog opn 
      Left            =   7680
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   4440
      TabIndex        =   7
      Top             =   4800
      Width           =   255
   End
   Begin VB.TextBox txtEmailAddress 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   4440
      TabIndex        =   3
      Top             =   2880
      Width           =   3135
   End
   Begin VB.TextBox txtMsg 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   4440
      TabIndex        =   2
      Top             =   3600
      Width           =   3135
   End
   Begin VB.TextBox txtTimes 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   4440
      TabIndex        =   1
      Top             =   4320
      Width           =   3135
   End
   Begin VB.CommandButton send 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Send"
      Height          =   255
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton send2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Send"
      Height          =   255
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton nxt 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Other"
      Height          =   255
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   89
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton turnoverform2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Your Scrolling Msgs"
      Height          =   255
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5280
      Width           =   1815
   End
   Begin VB.CommandButton turnoverform 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Default"
      Height          =   255
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Frame sf1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Scrolling Default Messages"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   4815
      Left            =   360
      TabIndex        =   10
      Top             =   360
      Width           =   3855
      Begin VB.CommandButton engflag 
         BackColor       =   &H00E0E0E0&
         Caption         =   "EnglandFlag"
         Height          =   195
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   96
         Top             =   4440
         Width           =   1215
      End
      Begin VB.CommandButton coffiecup 
         BackColor       =   &H00E0E0E0&
         Caption         =   "CoffieCup"
         Height          =   195
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   95
         Top             =   4440
         Width           =   1215
      End
      Begin VB.CommandButton spaceshuttle 
         BackColor       =   &H00E0E0E0&
         Caption         =   "SpaceShuttle"
         Height          =   195
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   4200
         Width           =   1215
      End
      Begin VB.CommandButton downarrow 
         BackColor       =   &H00E0E0E0&
         Caption         =   "DownArrow"
         Height          =   195
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   93
         Top             =   3960
         Width           =   1215
      End
      Begin VB.CommandButton star 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Star"
         Height          =   195
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   92
         Top             =   4200
         Width           =   1215
      End
      Begin VB.CommandButton starr 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Smile"
         Height          =   195
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   3960
         Width           =   1215
      End
      Begin VB.CommandButton house 
         BackColor       =   &H00E0E0E0&
         Caption         =   "House"
         Height          =   195
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Feedback"
         Height          =   195
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   3480
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cross"
         Height          =   195
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   3240
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Instructions"
         Height          =   195
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   3480
         Width           =   975
      End
      Begin VB.CommandButton dimond 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Dimond"
         Height          =   195
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CommandButton flood007 
         BackColor       =   &H00E0E0E0&
         Caption         =   "007"
         Height          =   195
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton msner 
         BackColor       =   &H00E0E0E0&
         Caption         =   "MSN"
         Height          =   195
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton tree 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tree"
         Height          =   195
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton Note 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Note"
         Height          =   195
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton xmark 
         BackColor       =   &H00E0E0E0&
         Caption         =   "XMark"
         Height          =   195
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton zshape 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ZShape"
         Height          =   195
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox custom 
         BackColor       =   &H00E0E0E0&
         Height          =   2805
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   600
         Width           =   2295
      End
      Begin VB.CommandButton sendme 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Send"
         Height          =   195
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   3480
         Width           =   855
      End
      Begin VB.CommandButton dvl 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Devil"
         Height          =   195
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CommandButton sk 
         BackColor       =   &H00E0E0E0&
         Caption         =   "DevilStick"
         Height          =   195
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CommandButton key 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Key"
         Height          =   195
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2520
         Width           =   1215
      End
      Begin VB.CommandButton clock 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Clock"
         Height          =   195
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label20 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Custom Flood"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   1440
         TabIndex        =   63
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Emotion Pics"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame othr 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Other Flooding Tools"
      Height          =   4815
      Left            =   360
      TabIndex        =   76
      Top             =   360
      Width           =   3855
      Begin VB.CommandButton ff1 
         Caption         =   "Font Flood"
         Height          =   255
         Left            =   1200
         TabIndex        =   86
         Top             =   4080
         Width           =   1695
      End
      Begin VB.TextBox f4 
         Height          =   285
         Left            =   600
         TabIndex        =   85
         Top             =   3720
         Width           =   2895
      End
      Begin VB.TextBox f2 
         Height          =   285
         Left            =   600
         TabIndex        =   84
         Top             =   3000
         Width           =   2895
      End
      Begin VB.TextBox f3 
         Height          =   285
         Left            =   600
         TabIndex        =   83
         Top             =   3360
         Width           =   2895
      End
      Begin VB.TextBox f1 
         Height          =   285
         Left            =   600
         TabIndex        =   82
         Top             =   2640
         Width           =   2895
      End
      Begin VB.CommandButton rf1 
         Caption         =   "Rainbow Flood"
         Height          =   255
         Left            =   1200
         TabIndex        =   81
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox r4 
         Height          =   285
         Left            =   600
         TabIndex        =   80
         Top             =   1560
         Width           =   2895
      End
      Begin VB.TextBox r2 
         Height          =   285
         Left            =   600
         TabIndex        =   79
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox r3 
         Height          =   285
         Left            =   600
         TabIndex        =   78
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox r1 
         Height          =   285
         Left            =   600
         TabIndex        =   77
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label12 
         Caption         =   "Font Flood"
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
         Left            =   120
         TabIndex        =   88
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Rainbow Flood"
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
         Left            =   120
         TabIndex        =   87
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame sf2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Your Scrolling Msg's"
      Height          =   4815
      Left            =   360
      TabIndex        =   24
      Top             =   360
      Width           =   3855
      Begin VB.TextBox se 
         Height          =   2535
         Left            =   1200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   58
         Text            =   "3.frx":A402E
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox cption 
         Height          =   285
         Left            =   1200
         TabIndex        =   57
         Top             =   3360
         Width           =   2535
      End
      Begin VB.CommandButton saveit 
         Caption         =   "Save"
         Height          =   195
         Left            =   1920
         TabIndex        =   56
         Top             =   4200
         Width           =   1815
      End
      Begin VB.OptionButton sm9 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Option1"
         Height          =   195
         Left            =   1680
         TabIndex        =   51
         Top             =   4200
         Width           =   255
      End
      Begin VB.OptionButton sm8 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Option1"
         Height          =   195
         Left            =   1320
         TabIndex        =   50
         Top             =   4200
         Width           =   255
      End
      Begin VB.OptionButton sm7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Option1"
         Height          =   195
         Left            =   3480
         TabIndex        =   49
         Top             =   3960
         Width           =   255
      End
      Begin VB.OptionButton sm6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Option1"
         Height          =   195
         Left            =   3120
         TabIndex        =   43
         Top             =   3960
         Width           =   255
      End
      Begin VB.OptionButton sm5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Option1"
         Height          =   195
         Left            =   2760
         TabIndex        =   42
         Top             =   3960
         Width           =   255
      End
      Begin VB.OptionButton sm4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Option1"
         Height          =   195
         Left            =   2400
         TabIndex        =   41
         Top             =   3960
         Width           =   255
      End
      Begin VB.OptionButton sm2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Option1"
         Height          =   195
         Left            =   1680
         TabIndex        =   40
         Top             =   3960
         Width           =   255
      End
      Begin VB.OptionButton sm3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Option1"
         Height          =   195
         Left            =   2040
         TabIndex        =   39
         Top             =   3960
         Width           =   255
      End
      Begin VB.OptionButton sm1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Option1"
         Height          =   195
         Left            =   1320
         TabIndex        =   38
         Top             =   3960
         Width           =   255
      End
      Begin VB.CommandButton openit 
         Caption         =   "Open"
         Height          =   255
         Left            =   2640
         TabIndex        =   36
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CommandButton ur8 
         Caption         =   "*8*"
         Height          =   375
         Left            =   120
         TabIndex        =   35
         Top             =   3600
         Width           =   975
      End
      Begin VB.CommandButton ur9 
         Caption         =   "*9*"
         Height          =   375
         Left            =   120
         TabIndex        =   34
         Top             =   4080
         Width           =   975
      End
      Begin VB.CommandButton ur7 
         Caption         =   "*7*"
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   3120
         Width           =   975
      End
      Begin VB.CommandButton ur5 
         Caption         =   "*5*"
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   2160
         Width           =   975
      End
      Begin VB.CommandButton ur6 
         Caption         =   "*6*"
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   2640
         Width           =   975
      End
      Begin VB.CommandButton ur4 
         Caption         =   "*4*"
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton ur2 
         Caption         =   "*2*"
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton ur3 
         Caption         =   "*3*"
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton ur1 
         Caption         =   "*1*"
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   975
      End
      Begin VB.Label stus 
         BackStyle       =   0  'Transparent
         Caption         =   "Status: "
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   4560
         Width           =   2055
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Caption:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   1200
         TabIndex        =   59
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   1560
         TabIndex        =   55
         Top             =   4200
         Width           =   135
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   1200
         TabIndex        =   54
         Top             =   4200
         Width           =   135
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   3360
         TabIndex        =   53
         Top             =   3960
         Width           =   135
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   3000
         TabIndex        =   52
         Top             =   3960
         Width           =   135
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   2640
         TabIndex        =   48
         Top             =   3960
         Width           =   135
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   2280
         TabIndex        =   47
         Top             =   3960
         Width           =   135
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   1920
         TabIndex        =   46
         Top             =   3960
         Width           =   135
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   1560
         TabIndex        =   45
         Top             =   3960
         Width           =   135
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   1200
         TabIndex        =   44
         Top             =   3960
         Width           =   135
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Save To:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   1200
         TabIndex        =   37
         Top             =   3720
         Width           =   975
      End
   End
   Begin VB.Image Image1 
      Height          =   1905
      Left            =   4440
      Picture         =   "3.frx":A403C
      Top             =   480
      Width           =   3630
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "r"
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
      TabIndex        =   62
      Top             =   840
      Width           =   1695
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   4320
      X2              =   4320
      Y1              =   360
      Y2              =   5520
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Invisible?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   4800
      TabIndex        =   9
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Times:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   4440
      TabIndex        =   6
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Victim's Email:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   4440
      TabIndex        =   5
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Message:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   4440
      TabIndex        =   4
      Top             =   3360
      Width           =   1455
   End
End
Attribute VB_Name = "scroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents Msn As MsgrObject
Attribute Msn.VB_VarHelpID = -1
Dim aUser As IMsgrUser
Dim Header As String
Private Sub Check1_Click()
If Check1.Value = 1 Then
txtEmailAddress.Visible = True
Label2.Visible = True
send2.Visible = False
send.Visible = True
End If
If Check1.Value = 0 Then
txtEmailAddress.Visible = False
Label2.Visible = False
send2.Visible = True
send.Visible = False
End If
End Sub
Private Sub clock_Click()
txtMsg.Text = "(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e):):):):)(e)(e)(e)(e)(e)(e)(e)(e)(e):)(i)(i)(i)(i):)(e)(e)(e)(e)(e)(e)(e)(e):)(i)(o)(o)(i):)(e)(e)(e)(e)(e)(e)(e)(e):)(i)(i)(i)(i):)(e)(e)(e)(e)(e)(e)(e)(e)(e):):):):)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e):):):):)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e):):):):)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e):):):):)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e):):):):)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e):):):):)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e):):):):)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)"
End Sub

Private Sub coffiecup_Click()
txtMsg.Text = "             (%)         (%)               (%)         (%)            (%)         (%)   (c)(c)(c)(c)(c)(c)(c)(c)(c)(c)(c)(c)      (c)(c)(c)(c)(c)(c)(c)(c)(c)(c)(o)(o)(o)      (c)(c)(c)(c)(c)(c)(c)(c)(c)(c)(o)   (o)         (c)(c)(c)(c)(c)(c)(c)(c)(c)(o)   (o)            (c)(c)(c)(c)(c)(c)(c)(c)(o)(o)(o)               (c)(c)(c)(c)(c)(c)"
End Sub

Private Sub Command1_Click()
order.Show
End Sub

Private Sub Command2_Click()
txtMsg.Text = "(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(*)(*)(z)(z)(z)(z)(z)(z)(z)(z)(*)(*)(z)(z)(z)(z)(z)(z)(*)(*)(*)(*)(*)(*)(z)(z)(z)(z)(*)(*)(*)(*)(*)(*)(z)(z)(z)(z)(z)(z)(*)(*)(z)(z)(z)(z)(z)(z)(z)(z)(*)(*)(z)(z)(z)(z)(z)(z)(z)(z)(*)(*)(z)(z)(z)(z)(z)(z)(z)(z)(*)(*)(z)(z)(z)(z)(z)(z)(z)(z)(*)(*)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)"
End Sub

Private Sub Command3_Click()
txtMsg.Text = "(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(z)(z)(*)(l)(l)(l)(l)(l)(l)(l)(l)(l)(l)(*)(z)(z)(*)(l)(8)(8)(8)(8)(8)(8)(8)(8)(l)(*)(z)(z)(*)(l)(8)(e)(e)(e)(e)(e)(e)(8)(l)(*)(z)(z)(*)(l)(8)(e)(b)(b)(b)(b)(e)(8)(l)(*)(z)(z)(*)(l)(8)(e)(b)(b)(b)(b)(e)(8)(l)(*)(z)(z)(*)(l)(8)(e)(b)(b)(b)(b)(e)(8)(l)(*)(z)(z)(*)(l)(8)(e)(e)(e)(e)(e)(e)(8)(l)(*)(z)(z)(*)(l)(8)(8)(8)(8)(8)(8)(8)(8)(l)(*)(z)(z)(*)(l)(l)(l)(l)(l)(l)(l)(l)(l)(l)(*)(z)(z)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)"
End Sub


Private Sub downarrow_Click()
txtMsg.Text = "(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(e)(e)(e)(e)(*)(*)(*)(*)(*)(*)(*)(*)(e)(e)(e)(e)(*)(*)(*)(*)(*)(*)(*)(*)(e)(e)(e)(e)(*)(*)(*)(*)(*)(*)(*)(*)(e)(e)(e)(e)(*)(*)(*)(*)(*)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(*)(*)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(*)(*)(*)(e)(e)(e)(e)(e)(e)(e)(e)(*)(*)(*)(*)(*)(e)(e)(e)(e)(e)(e)(*)(*)(*)(*)(*)(*)(*)(e)(e)(e)(e)(*)(*)(*)(*)(*)(*)(*)(*)(*)(e)(e)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)"
End Sub

Private Sub engflag_Click()
txtMsg.Text = " :[(o)(o)(o)(o)(o)(o)(o)(o)(o):[(o)(o)(o)(o):@(o)(o)(o)(o):[(o)(o)(o)(o):@(o)(o)(o)(o):[(o)(o)(o)(o):@(o)(o)(o)(o):[(o):@:@:@:@:@:@:@(o):[(o)(o)(o)(o):@(o)(o)(o)(o):[(o)(o)(o)(o):@(o)(o)(o)(o):[(o)(o)(o)(o):@(o)(o)(o)(o):[:[:[:[:["
End Sub

Private Sub spaceshuttle_Click()
txtMsg.Text = ":[:[:[:[:[:[:[:[:[:[:[:[:[(6)(6)(6)(6)(6):[:[:[(6)(6)(6)(6)(6):[(6)(6)(6)(6):[:[:[(6)(6)(6)(6):[:[:[(6)(6)(6):[:[:[(6)(6)(6):[:[:[:[:[(*)(*):[:[:[(*)(*):[:[:[:[:[:[(*)(*)(*)(*)(*)(*)(*):[:[:[:[:[:[(*)(*):[:[:[(*)(*):[:[:[:[:[:[(*)(*):[:[:[(*)(*):[:[:[:[:[:[(*)(*)(*)(*)(*)(*)(*):[:[:[:[:[:[(*)(*)(*)(*)(*)(*)(*):[:[:[:[:[:[:[(*)(*)(*)(*)(*):[:[:[:[:[:[:[:[:[(*)(*)(*):[:[:[:[:[:[:[:[:[:[:[(*):[:[:[:[:[:[:[:[:[:[:[:[:[:[:[:[:[:[:["
End Sub

Private Sub star_Click()
txtMsg.Text = "(S)(S)(S)(S)(S)(*)(S)(S)(S)(S)(S)(S)(S)(S)(S)(*)(*)(*)(S)(S)(S)(S)(S)(S)(S)(*)(*)(*)(*)(*)(S)(S)(S)(S)(S)(*)(*)(*)(*)(*)(*)(*)(S)(S)(S)(*)(*)(*)(*)(*)(*)(*)(*)(*)(S)(S)(S)(*)(*)(*)(*)(*)(*)(*)(S)(S)(S)(S)(S)(*)(*)(*)(*)(*)(S)(S)(S)(S)(S)(S)(S)(*)(*)(*)(S)(S)(S)(S)(S)(S)(S)(S)(S)(*)(S)(S)(S)(S)(S)"
End Sub

Private Sub starr_Click()
txtMsg.Text = "(d)(d)(d)(d)(d)(d)(d)(d)(d)(d)(d)(d)(d):):):)(d)(d)(d)(d):):):)(d)(d):):):)(d)(d)(d)(d):):):)(d)(d):):):)(d)(d)(d)(d):):):)(d)(d)(d)(d)(d)(d):):)(d)(d)(d)(d)(d)(d)(d)(d)(d)(d):):)(d)(d)(d)(d)(d)(d):)(d)(d)(d)(d)(d)(d)(d)(d):)(d)(d):((d)(d)(d)(d)(d)(d)(d)(d):)(d)(d):):):):):):):):):):)(d)(d)(d)(d)(d)(d)(d)(d)(d)(d)(d)(d)(d)"
End Sub

Private Sub dimond_Click()
txtMsg.Text = "(l)(l)(l)(l)(l)(l)(l)(l)(l)(l)(l)(l)(l)(l)(l)(l)(*)(l)(l)(l)(l)(l)(l)(l)(l)(l)(*)(*)(*)(l)(l)(l)(l)(l)(l)(l)(*)(*)(*)(*)(*)(l)(l)(l)(l)(l)(*)(*)(*)(*)(*)(*)(*)(l)(l)(l)(l)(l)(*)(*)(*)(*)(*)(l)(l)(l)(l)(l)(l)(l)(*)(*)(*)(l)(l)(l)(l)(l)(l)(l)(l)(l)(*)(l)(l)(l)(l)(l)(l)(l)(l)(l)(l)(l)(l)(l)(l)(l)(l)"
End Sub

Private Sub dvl_Click()
txtMsg.Text = ":|:|:|:|:|:|:|:|:|(6)(6):|:|(6)(6):|:|(6):|:|:|:|(6):|:|(a)(6)(6)(6)(6)(6):|:|(6)(6)(6)(6)(6)(6):|:|(6)(o)(6)(6)(o)(6):|:|(6)(6)(6)(6)(6)(6):|:|:|(6)(*)(*)(6):|:|:|:|(6)(6)(6)(6):|:|:|:|:|(6)(6):|:|:|"
End Sub

Private Sub ff1_Click()
Dim count As Integer
count = 0
While count < txtTimes.Text
    Msn.CreateUser(txtEmail.Text, Msn.Services.PrimaryService).SendText Headr, r1.Text, MMSGTYPE_ALL_RESULTS
    Msn.CreateUser(txtEmail.Text, Msn.Services.PrimaryService).SendText Headr1, r2.Text, MMSGTYPE_ALL_RESULTS
    Msn.CreateUser(txtEmail.Text, Msn.Services.PrimaryService).SendText Headr2, r3.Text, MMSGTYPE_ALL_RESULTS
    Msn.CreateUser(txtEmail.Text, Msn.Services.PrimaryService).SendText Headr3, r4.Text, MMSGTYPE_ALL_RESULTS
   count = count + 1
Wend
End Sub

Private Sub house_Click()
txtMsg.Text = "                      (6)   (6)(6)(6)(6)(6)(6)(6)(6)(6)(6)(6)(6)(6)(6)(6)(6)(6)(6)(a)(a)(a)(a)(a)(a)(a)(a)(a)(a)(a)(g)(g)(a)(a)(a)(a)(g)(g)(a)(a)(g)(g)(a)(a)(a)(a)(g)(g)(a)(a)(a)(a)(a)(a)(a)(a)(a)(a)(a)(a)(a)(a)(a)(a)(a)(a)(a)(a)(a)(a)(g)(g)(a)(#)(#)(a)(g)(g)(a)(a)(g)(g)(a)(#)(#)(a)(g)(g)(a)(a)(a)(a)(a)(#)(#)(a)(a)(a)(a)"
End Sub

Private Sub key_Click()
txtMsg.Text = "(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e):):):):)(e)(e)(e)(e)(e)(e)(e)(e)(e):)(e)(e)(e)(e):)(e)(e)(e)(e)(e)(e)(e)(e):)(e)(e)(e)(e):)(e)(e)(e)(e)(e)(e)(e)(e):)(e)(e)(e)(e):)(e)(e)(e)(e)(e)(e)(e)(e)(e):):):):)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e):):)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e):):):):)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e):):)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e):):):):)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e):):)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e):):):):)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)"
End Sub



Private Sub msner_Click()
txtMsg.Text = ""
txtMsg.Text = "(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(8)(*)(*)(*)(8)(*)(8)(8)(8)(8)(*)(8)(8)(8)(8)(*)(*)(*)(8)(8)(*)(8)(8)(*)(8)(*)(*)(*)(*)(8)(*)(*)(8)(*)(*)(*)(8)(*)(8)(*)(8)(*)(8)(8)(8)(8)(*)(8)(*)(*)(8)(*)(*)(*)(8)(*)(*)(*)(8)(*)(*)(*)(*)(8)(*)(8)(*)(*)(8)(*)(*)(*)(8)(*)(*)(*)(8)(*)(8)(8)(8)(8)(*)(8)(*)(*)(8)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)"
End Sub

Private Sub note_Click()
txtMsg.Text = ""
txtMsg.Text = "(i)(i)(i)(i)(i)(8)(8)(i)(i)(i)(i)(i)(i)(i)(i)(i)(i)(8)(i)(8)(i)(i)(i)(i)(i)(i)(i)(i)(i)(8)(i)(i)(8)(i)(i)(i)(i)(i)(i)(i)(i)(8)(i)(i)(i)(8)(i)(i)(i)(i)(i)(i)(i)(8)(i)(i)(i)(i)(i)(i)(i)(i)(i)(i)(8)(8)(8)(i)(i)(i)(i)(i)(i)(i)(i)(8)(8)(8)(8)(8)(i)(i)(i)(i)(i)(i)(8)(8)(8)(8)(8)(8)(i)(i)(i)(i)(i)(i)(i)(8)(8)(8)(8)(8)(i)(i)(i)(i)(i)(i)(i)(i)(8)(8)(8)(8)(i)(i)(i)(i)(i)(i)(i)(i)(i)(i)(i)(i)(i)(i)(i)(i)"
End Sub

Private Sub nxt_Click()
sf2.Visible = True
sf1.Visible = False
othr.Visible = False
turnoverform2.Visible = False
turnoverform.Visible = True
nxt.Visible = False
End Sub

Private Sub openit_Click()
On Error GoTo 10
'On Error Goto Command For Cancel In The Dialog Box
'On Error GoTo 10
'CommonDialog1.ShowOpen Renders The Open DialogBox
opn.ShowOpen
'-------------------
'CommonDialog1.FileName Is The Value Of The Path & Url The User Selected
'The Open Command Loads A File Into RAM[ For Input As (#1(FileMarker While Open)).You Can Use #(AnyNumber) Consistantly Through The Sub
Open opn.FileName For Input As #1
'-------------------
'The Input Statement Extracts Information From A File From The Same Order You Saved It In
'It Is Important to Extract That Information In The Same Order You Saved It In
'To input More Inforemation: Ex: Input #1, VariableA, VariableB,VariableC
'Text1.Text = VariableA: Text2.Text = VariableB: Text3.Text = Button1.Caption   etc...
Input #1, AnyVariableWillDo$
se.Text = ""
'-------------------
'Apply The Extracted Information To The TextBox
se.Text = AnyVariableWillDo$
'-------------------
'Refresh The RichTextBox
se.Refresh
'-------------------
'It Is Important To Close Open Files At The End of The Sub, Otherwise Errors Will Occur
Close #1

'-------------------
'Now That We Have Opened A File We Can Make The Save Button Visible
'Apply The File Name To The StatusBar
10
End Sub





Private Sub rf1_Click()
Dim count As Integer
count = 0
While count < txtTimes.Text
    Msn.CreateUser(txtEmail.Text, Msn.Services.PrimaryService).SendText Header0, r1.Text, MMSGTYPE_ALL_RESULTS
    Msn.CreateUser(txtEmail.Text, Msn.Services.PrimaryService).SendText Header1, r2.Text, MMSGTYPE_ALL_RESULTS
    Msn.CreateUser(txtEmail.Text, Msn.Services.PrimaryService).SendText Header2, r3.Text, MMSGTYPE_ALL_RESULTS
    Msn.CreateUser(txtEmail.Text, Msn.Services.PrimaryService).SendText Header3, r4.Text, MMSGTYPE_ALL_RESULTS
   count = count + 1
Wend
End Sub
Private Sub saveit_Click()
If sm1.Value = True Then
CreateKey "HKCU\Software\TechnoBOT\Scrolling\usr1", se.Text
CreateKey "HKCU\Software\TechnoBOT\Scrolling\cpt1", cption.Text
stus.Caption = "Status: Saved To Button 1"
ur1.Caption = ReadKey("HKCU\Software\TechnoBOT\Scrolling\cpt1")
End If
If sm2.Value = True Then
CreateKey "HKCU\Software\TechnoBOT\Scrolling\usr2", se.Text
CreateKey "HKCU\Software\TechnoBOT\Scrolling\cpt2", cption.Text
stus.Caption = "Status: Saved To Button 2"
ur2.Caption = ReadKey("HKCU\Software\TechnoBOT\Scrolling\cpt2")
End If
If sm3.Value = True Then
CreateKey "HKCU\Software\TechnoBOT\Scrolling\usr3", se.Text
CreateKey "HKCU\Software\TechnoBOT\Scrolling\cpt3", cption.Text
stus.Caption = "Status: Saved To Button 3"
ur3.Caption = ReadKey("HKCU\Software\TechnoBOT\Scrolling\cpt3")
End If
If sm4.Value = True Then
CreateKey "HKCU\Software\TechnoBOT\Scrolling\usr4", se.Text
CreateKey "HKCU\Software\TechnoBOT\Scrolling\cpt4", cption.Text
stus.Caption = "Status: Saved To Button 4"
ur4.Caption = ReadKey("HKCU\Software\TechnoBOT\Scrolling\cpt4")
End If
If sm5.Value = True Then
CreateKey "HKCU\Software\TechnoBOT\Scrolling\usr5", se.Text
CreateKey "HKCU\Software\TechnoBOT\Scrolling\cpt5", cption.Text
stus.Caption = "Status: Saved To Button 5"
ur5.Caption = ReadKey("HKCU\Software\TechnoBOT\Scrolling\cpt5")
End If
If sm6.Value = True Then
CreateKey "HKCU\Software\TechnoBOT\Scrolling\usr6", se.Text
CreateKey "HKCU\Software\TechnoBOT\Scrolling\cpt6", cption.Text
stus.Caption = "Status: Saved To Button 6"
ur6.Caption = ReadKey("HKCU\Software\TechnoBOT\Scrolling\cpt6")
End If
If sm7.Value = True Then
CreateKey "HKCU\Software\TechnoBOT\Scrolling\usr7", se.Text
CreateKey "HKCU\Software\TechnoBOT\Scrolling\cpt7", cption.Text
stus.Caption = "Status: Saved To Button 7"
ur7.Caption = ReadKey("HKCU\Software\TechnoBOT\Scrolling\cpt7")
End If
If sm8.Value = True Then
CreateKey "HKCU\Software\TechnoBOT\Scrolling\usr8", se.Text
CreateKey "HKCU\Software\TechnoBOT\Scrolling\cpt8", cption.Text
stus.Caption = "Status: Saved To Button 8"
ur8.Caption = ReadKey("HKCU\Software\TechnoBOT\Scrolling\cpt8")
End If
If sm9.Value = True Then
CreateKey "HKCU\Software\TechnoBOT\Scrolling\usr9", se.Text
CreateKey "HKCU\Software\TechnoBOT\Scrolling\cpt9", cption.Text
stus.Caption = "Status: Saved To Button 9"
ur9.Caption = ReadKey("HKCU\Software\TechnoBOT\Scrolling\cpt9")
End If
End Sub

Private Sub send_Click()
If txtMsg.Text = "" Then
MsgBox "Message Has Not Been Entered"
GoTo 1
End If
Dim times, x As Integer
Set aUser = Msn.CreateUser(txtEmailAddress.Text, Msn.Services.PrimaryService)
times = txtTimes.Text
On Error Resume Next
For x = 1 To times
aUser.SendText Header, txtMsg, MMSGTYPE_ALL_RESULTS
Next x
MsgBox ("All Messages Sent")
1:
End Sub

Private Sub flood007_Click()
txtMsg.Text = "(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(D)(D)(D)(D)(*)(D)(D)(D)(D)(*)(D)(D)(D)(D)(D)(*)(*)(D)(*)(*)(D)(*)(D)(*)(*)(D)(*)(*)(*)(*)(D)(*)(*)(*)(D)(*)(*)(D)(*)(D)(*)(*)(D)(*)(*)(*)(D)(*)(*)(*)(*)(D)(*)(*)(D)(*)(D)(*)(*)(D)(*)(*)(D)(*)(*)(*)(*)(*)(D)(D)(D)(D)(*)(D)(D)(D)(D)(*)(D)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)"
End Sub


Private Sub Form_Load()
Set Msn = New MsgrObject
Header = "MIME-Version: 1.0" & vbCrLf & "Content-Type: text/plain; charset=UTF-8" & vbCrLf & "X-MMS-IM-Format: FN=Comic%20Sans%20MS; EF=B; CO=ffa500; CS=0; PF=42"
txtEmailAddress.Visible = False
Label2.Visible = False
send.Visible = False
ur1.Caption = ReadKey("HKCU\Software\TechnoBOT\Scrolling\cpt1")
ur2.Caption = ReadKey("HKCU\Software\TechnoBOT\Scrolling\cpt2")
ur3.Caption = ReadKey("HKCU\Software\TechnoBOT\Scrolling\cpt3")
ur4.Caption = ReadKey("HKCU\Software\TechnoBOT\Scrolling\cpt4")
ur5.Caption = ReadKey("HKCU\Software\TechnoBOT\Scrolling\cpt5")
ur6.Caption = ReadKey("HKCU\Software\TechnoBOT\Scrolling\cpt6")
ur7.Caption = ReadKey("HKCU\Software\TechnoBOT\Scrolling\cpt7")
ur8.Caption = ReadKey("HKCU\Software\TechnoBOT\Scrolling\cpt8")
ur9.Caption = ReadKey("HKCU\Software\TechnoBOT\Scrolling\cpt9")
cust1.Text = ReadKey("HKCU\Software\TechnoBOT\Scrolling\usr1")
cust2.Text = ReadKey("HKCU\Software\TechnoBOT\Scrolling\usr2")
cust3.Text = ReadKey("HKCU\Software\TechnoBOT\Scrolling\usr3")
cust4.Text = ReadKey("HKCU\Software\TechnoBOT\Scrolling\usr4")
cust5.Text = ReadKey("HKCU\Software\TechnoBOT\Scrolling\usr5")
cust6.Text = ReadKey("HKCU\Software\TechnoBOT\Scrolling\usr6")
cust7.Text = ReadKey("HKCU\Software\TechnoBOT\Scrolling\usr7")
cust8.Text = ReadKey("HKCU\Software\TechnoBOT\Scrolling\usr8")
cust9.Text = ReadKey("HKCU\Software\TechnoBOT\Scrolling\usr9")
opn.Filter = "Text Files|*.TXT"
End Sub

Private Sub send2_Click()
On Error GoTo ER
Dim times, x As Integer
times = txtTimes.Text
For x = 1 To times
SendMessag txtMsg.Text
Next x
MsgBox ("All Messages Sent")
ER:
End Sub

Private Sub sendme_Click()
sit
End Sub
Sub sit()
SelectTxt custom
Dim spliter As Variant
spliter = Split(custom.Text, ",")
For i = 0 To UBound(spliter) - 1
custom.SelStart = Len(custom.Text)
Select Case spliter(i)
Case "coffiecup"
SendMessag "             (%)         (%)               (%)         (%)            (%)         (%)   (c)(c)(c)(c)(c)(c)(c)(c)(c)(c)(c)(c)      (c)(c)(c)(c)(c)(c)(c)(c)(c)(c)(o)(o)(o)      (c)(c)(c)(c)(c)(c)(c)(c)(c)(c)(o)   (o)         (c)(c)(c)(c)(c)(c)(c)(c)(c)(o)   (o)            (c)(c)(c)(c)(c)(c)(c)(c)(o)(o)(o)               (c)(c)(c)(c)(c)(c)"
Case "englandflag"
SendMessag " :[(o)(o)(o)(o)(o)(o)(o)(o)(o):[(o)(o)(o)(o):@(o)(o)(o)(o):[(o)(o)(o)(o):@(o)(o)(o)(o):[(o)(o)(o)(o):@(o)(o)(o)(o):[(o):@:@:@:@:@:@:@(o):[(o)(o)(o)(o):@(o)(o)(o)(o):[(o)(o)(o)(o):@(o)(o)(o)(o):[(o)(o)(o)(o):@(o)(o)(o)(o):[:[:[:[:["
Case "spaceshuttle"
SendMessag ":[:[:[:[:[:[:[:[:[:[:[:[:[(6)(6)(6)(6)(6):[:[:[(6)(6)(6)(6)(6):[(6)(6)(6)(6):[:[:[(6)(6)(6)(6):[:[:[(6)(6)(6):[:[:[(6)(6)(6):[:[:[:[:[(*)(*):[:[:[(*)(*):[:[:[:[:[:[(*)(*)(*)(*)(*)(*)(*):[:[:[:[:[:[(*)(*):[:[:[(*)(*):[:[:[:[:[:[(*)(*):[:[:[(*)(*):[:[:[:[:[:[(*)(*)(*)(*)(*)(*)(*):[:[:[:[:[:[(*)(*)(*)(*)(*)(*)(*):[:[:[:[:[:[:[(*)(*)(*)(*)(*):[:[:[:[:[:[:[:[:[(*)(*)(*):[:[:[:[:[:[:[:[:[:[:[(*):[:[:[:[:[:[:[:[:[:[:[:[:[:[:[:[:[:[:["
Case "downarrow"
SendMessag "(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(e)(e)(e)(e)(*)(*)(*)(*)(*)(*)(*)(*)(e)(e)(e)(e)(*)(*)(*)(*)(*)(*)(*)(*)(e)(e)(e)(e)(*)(*)(*)(*)(*)(*)(*)(*)(e)(e)(e)(e)(*)(*)(*)(*)(*)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(*)(*)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(*)(*)(*)(e)(e)(e)(e)(e)(e)(e)(e)(*)(*)(*)(*)(*)(e)(e)(e)(e)(e)(e)(*)(*)(*)(*)(*)(*)(*)(e)(e)(e)(e)(*)(*)(*)(*)(*)(*)(*)(*)(*)(e)(e)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)"
Case "star"
SendMessag "(S)(S)(S)(S)(S)(*)(S)(S)(S)(S)(S)(S)(S)(S)(S)(*)(*)(*)(S)(S)(S)(S)(S)(S)(S)(*)(*)(*)(*)(*)(S)(S)(S)(S)(S)(*)(*)(*)(*)(*)(*)(*)(S)(S)(S)(*)(*)(*)(*)(*)(*)(*)(*)(*)(S)(S)(S)(*)(*)(*)(*)(*)(*)(*)(S)(S)(S)(S)(S)(*)(*)(*)(*)(*)(S)(S)(S)(S)(S)(S)(S)(*)(*)(*)(S)(S)(S)(S)(S)(S)(S)(S)(S)(*)(S)(S)(S)(S)(S)"
Case "smile"
SendMessag "(d)(d)(d)(d)(d)(d)(d)(d)(d)(d)(d)(d)(d):):):)(d)(d)(d)(d):):):)(d)(d):):):)(d)(d)(d)(d):):):)(d)(d):):):)(d)(d)(d)(d):):):)(d)(d)(d)(d)(d)(d):):)(d)(d)(d)(d)(d)(d)(d)(d)(d)(d):):)(d)(d)(d)(d)(d)(d):)(d)(d)(d)(d)(d)(d)(d)(d):)(d)(d):((d)(d)(d)(d)(d)(d)(d)(d):)(d)(d):):):):):):):):):):)(d)(d)(d)(d)(d)(d)(d)(d)(d)(d)(d)(d)(d)"
Case "house"
SendMessag "                      (6)   (6)(6)(6)(6)(6)(6)(6)(6)(6)(6)(6)(6)(6)(6)(6)(6)(6)(6)(a)(a)(a)(a)(a)(a)(a)(a)(a)(a)(a)(g)(g)(a)(a)(a)(a)(g)(g)(a)(a)(g)(g)(a)(a)(a)(a)(g)(g)(a)(a)(a)(a)(a)(a)(a)(a)(a)(a)(a)(a)(a)(a)(a)(a)(a)(a)(a)(a)(a)(a)(g)(g)(a)(#)(#)(a)(g)(g)(a)(a)(g)(g)(a)(#)(#)(a)(g)(g)(a)(a)(a)(a)(a)(#)(#)(a)(a)(a)(a)"
Case "feedback"
SendMessag "(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(z)(z)(*)(l)(l)(l)(l)(l)(l)(l)(l)(l)(l)(*)(z)(z)(*)(l)(8)(8)(8)(8)(8)(8)(8)(8)(l)(*)(z)(z)(*)(l)(8)(e)(e)(e)(e)(e)(e)(8)(l)(*)(z)(z)(*)(l)(8)(e)(b)(b)(b)(b)(e)(8)(l)(*)(z)(z)(*)(l)(8)(e)(b)(b)(b)(b)(e)(8)(l)(*)(z)(z)(*)(l)(8)(e)(b)(b)(b)(b)(e)(8)(l)(*)(z)(z)(*)(l)(8)(e)(e)(e)(e)(e)(e)(8)(l)(*)(z)(z)(*)(l)(8)(8)(8)(8)(8)(8)(8)(8)(l)(*)(z)(z)(*)(l)(l)(l)(l)(l)(l)(l)(l)(l)(l)(*)(z)(z)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)"
Case "cross"
SendMessag "(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(*)(*)(z)(z)(z)(z)(z)(z)(z)(z)(*)(*)(z)(z)(z)(z)(z)(z)(*)(*)(*)(*)(*)(*)(z)(z)(z)(z)(*)(*)(*)(*)(*)(*)(z)(z)(z)(z)(z)(z)(*)(*)(z)(z)(z)(z)(z)(z)(z)(z)(*)(*)(z)(z)(z)(z)(z)(z)(z)(z)(*)(*)(z)(z)(z)(z)(z)(z)(z)(z)(*)(*)(z)(z)(z)(z)(z)(z)(z)(z)(*)(*)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)"
Case "1"
SendMessag cust1.Text
Case "2"
SendMessag cust2.Text
Case "3"
SendMessag cust3.Text
Case "4"
SendMessag cust4.Text
Case "5"
SendMessag cust5.Text
Case "6"
SendMessag cust6.Text
Case "7"
SendMessag cust7.Text
Case "8"
SendMessag cust8.Text
Case "9"
SendMessag cust9.Text
Case "dimond"
SendMessag "(l)(l)(l)(l)(l)(l)(l)(l)(l)(l)(l)(l)(l)(l)(l)(l)(*)(l)(l)(l)(l)(l)(l)(l)(l)(l)(*)(*)(*)(l)(l)(l)(l)(l)(l)(l)(*)(*)(*)(*)(*)(l)(l)(l)(l)(l)(*)(*)(*)(*)(*)(*)(*)(l)(l)(l)(l)(l)(*)(*)(*)(*)(*)(l)(l)(l)(l)(l)(l)(l)(*)(*)(*)(l)(l)(l)(l)(l)(l)(l)(l)(l)(*)(l)(l)(l)(l)(l)(l)(l)(l)(l)(l)(l)(l)(l)(l)(l)(l)"
Case "clock"
SendMessag "(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e):):):):)(e)(e)(e)(e)(e)(e)(e)(e)(e):)(i)(i)(i)(i):)(e)(e)(e)(e)(e)(e)(e)(e):)(i)(o)(o)(i):)(e)(e)(e)(e)(e)(e)(e)(e):)(i)(i)(i)(i):)(e)(e)(e)(e)(e)(e)(e)(e)(e):):):):)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e):):):):)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e):):):):)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e):):):):)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e):):):):)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e):):):):)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e):):):):)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)"
Case "key"
SendMessag "(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e):):):):)(e)(e)(e)(e)(e)(e)(e)(e)(e):)(e)(e)(e)(e):)(e)(e)(e)(e)(e)(e)(e)(e):)(e)(e)(e)(e):)(e)(e)(e)(e)(e)(e)(e)(e):)(e)(e)(e)(e):)(e)(e)(e)(e)(e)(e)(e)(e)(e):):):):)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e):):)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e):):):):)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e):):)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e):):):):)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e):):)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e):):):):)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)(e)"
Case "devilstick"
txtMsg.Text = "(n)(n)(n)(n)(n)(n)(n)(n)(n)(n)(n)(n)(n)(6)(6)(n)(6)(6)(n)(6)(6)(n)(n)(n)(6)(6)(n)(6)(6)(n)(6)(6)(n)(n)(n)(6)(6)(n)(6)(6)(n)(6)(6)(n)(n)(n)(6)(6)(n)(6)(6)(n)(6)(6)(n)(n)(n)(n)(6)(n)(6)(6)(6)(6)(n)(n)(n)(n)(n)(6)(6)(6)(6)(6)(6)(n)(n)(n)(n)(n)(n)(n)(6)(6)(n)(n)(n)(n)(n)(n)(n)(n)(n)(6)(6)(n)(n)(n)(n)(n)(n)(n)(n)(n)(6)(6)(n)(n)(n)(n)(n)(n)(n)(n)(n)(6)(6)(n)(n)(n)(n)(n)(n)(n)(n)(n)(6)(6)(n)(n)(n)(n)(n)(n)(n)(n)(n)(n)(n)(n)(n)(n)(n)"
Case "devil"
SendMessag ":|:|:|:|:|:|:|:|:|(6)(6):|:|(6)(6):|:|(6):|:|:|:|(6):|:|(a)(6)(6)(6)(6)(6):|:|(6)(6)(6)(6)(6)(6):|:|(6)(o)(6)(6)(o)(6):|:|(6)(6)(6)(6)(6)(6):|:|:|(6)(*)(*)(6):|:|:|:|(6)(6)(6)(6):|:|:|:|:|(6)(6):|:|:|"
Case "zshape"
SendMessag "(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(m)(m)(8)(8)(e)(e)(e)(*)(*)(m)(m)(m)(8)(8)(e)(e)(*)(*)(m)(m)(m)(m)(8)(8)(e)(*)(*)(m)(m)(m)(8)(8)(e)(e)(*)(*)(m)(m)(8)(8)(e)(e)(e)(*)(*)(m)(8)(8)(e)(e)(e)(e)(*)(*)(m)(m)(8)(8)(e)(e)(e)(*)(*)(m)(m)(m)(8)(8)(e)(e)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)"
Case "xmark"
SendMessag "(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z):|:|:|:|:|:|:|:|:|:|:|(z)(z):|:@:@:|:|:|:|:|:@:@:|(z)(z):|:|:@:@:|:|:|:@:@:|:|(z)(z):|:|:|:@:@:|:@:@:|:|:|(z)(z):|:|:|:|:@:@:@:|:|:|:|(z)(z):|:|:|:@:@:|:@:@:|:|:|(z)(z):|:|:@:@:|:|:|:@:@:|:|(z)(z):|:@:@:|:|:|:|:|:@:@:|(z)(z):|:|:|:|:|:|:|:|:|:|:|(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)"
Case "note"
SendMessag "(i)(i)(i)(i)(i)(8)(8)(i)(i)(i)(i)(i)(i)(i)(i)(i)(i)(8)(i)(8)(i)(i)(i)(i)(i)(i)(i)(i)(i)(8)(i)(i)(8)(i)(i)(i)(i)(i)(i)(i)(i)(8)(i)(i)(i)(8)(i)(i)(i)(i)(i)(i)(i)(8)(i)(i)(i)(i)(i)(i)(i)(i)(i)(i)(8)(8)(8)(i)(i)(i)(i)(i)(i)(i)(i)(8)(8)(8)(8)(8)(i)(i)(i)(i)(i)(i)(8)(8)(8)(8)(8)(8)(i)(i)(i)(i)(i)(i)(i)(8)(8)(8)(8)(8)(i)(i)(i)(i)(i)(i)(i)(i)(8)(8)(8)(8)(i)(i)(i)(i)(i)(i)(i)(i)(i)(i)(i)(i)(i)(i)(i)(i)"
Case "tree"
SendMessag "(m)(m)(m)(m)(m)(*)(m)(m)(m)(m)(m)(m)(m)(m)(m)(*)(*)(*)(m)(m)(m)(m)(m)(m)(m)(*)(*)(*)(*)(*)(m)(m)(m)(m)(m)(m)(m)(*)(*)(*)(m)(m)(m)(m)(m)(m)(m)(*)(*)(*)(*)(*)(m)(m)(m)(m)(m)(*)(*)(*)(*)(*)(*)(*)(m)(m)(m)(m)(m)(m)(b)(b)(b)(m)(m)(m)(m)"
Case "a"
SendMessag "(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@:@:@:@(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h):@:@:@:@(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)"
Case "b"
SendMessag "(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@:@:@:@(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h):@:@:@:@(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h):@:@:@:@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)"
Case "c"
SendMessag "(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@:@:@:@(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@:@:@:@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)"
Case "d"
SendMessag "(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@:@:@:@(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h):@:@:@:@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)"
Case "e"
SendMessag "(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@:@:@:@(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@:@:@:@(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@:@:@:@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)"
Case "f"
SendMessag "(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@:@:@:@(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@:@:@(h)(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)"
Case "g"
SendMessag "(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@:@:@:@(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h):@:@(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h):@:@:@:@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)"
Case "h"
SendMessag "(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h):@:@:@:@:@(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)"
Case "i"
SendMessag "(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@:@:@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@:@:@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)"
Case "j"
SendMessag "(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@:@:@:@:@(h)(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h):@(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h):@(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@:@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)"
Case "k"
SendMessag "(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h):@(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h):@(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h):@:@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)"
Case "l"
SendMessag "(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@:@:@:@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)"
Case "m"
SendMessag "(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h):@(h)(h)(h)(h)(h)(h):@:@(h):@:@(h)(h)(h)(h)(h)(h):@(h):@(h):@(h)(h)(h)(h)(h)(h):@(h)(h)(h):@(h)(h)(h)(h)(h)(h):@(h)(h)(h):@(h)(h)(h)(h)(h)(h):@(h)(h)(h):@(h)(h)(h)(h)(h)(h):@(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)"
Case "n"
SendMessag "(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@:@:@:@(h)(h)(h)(h)(h)(h):@(h)(h)(h):@(h)(h)(h)(h)(h):@(h)(h)(h):@(h)(h)(h)(h)(h)(h):@(h)(h)(h):@(h)(h)(h)(h)(h)(h):@(h)(h)(h):@(h)(h)(h)(h)(h)(h):@(h)(h)(h):@(h)(h)(h)(h)(h)(h):@(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)"
Case "o"
SendMessag "(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@:@:@:@:@(h)(h)(h)(h)(h)(h):@(h)(h)(h):@(h)(h)(h)(h)(h)(h):@(h)(h)(h):@(h)(h)(h)(h)(h)(h):@(h)(h)(h):@(h)(h)(h)(h)(h)(h):@(h)(h)(h):@(h)(h)(h)(h)(h)(h):@(h)(h)(h):@(h)(h)(h)(h)(h)(h):@:@:@:@:@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)"
Case "p"
SendMessag "(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@:@:@:@:@(h)(h)(h)(h)(h)(h):@(h)(h)(h)(h):@(h)(h)(h)(h)(h):@(h)(h)(h)(h):@(h)(h)(h)(h)(h):@:@:@:@:@(h)(h)(h)(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)"
Case "q"
SendMessag "(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@:@:@:@:@(h)(h)(h)(h)(h)(h):@(h)(h)(h)(h):@(h)(h)(h)(h)(h):@(h)(h)(h)(h):@(h)(h)(h)(h)(h):@(h)(h)(h)(h):@(h)(h)(h)(h)(h):@(h)(h)(h):@:@(h)(h)(h)(h)(h):@(h)(h)(h)(h):@(h)(h)(h)(h)(h):@:@:@:@:@(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)"
Case "r"
SendMessag "(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@:@:@:@:@(h)(h)(h)(h)(h):@(h)(h)(h)(h):@(h)(h)(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)"
Case "s"
SendMessag "(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@:@:@:@:@(h)(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h):@:@:@:@:@(h)(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h)(h):@:@:@:@:@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)"
Case "t"
SendMessag "(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@:@:@:@:@(h)(h)(h)(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)"
Case "u"
SendMessag "(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h):@(h)(h)(h)(h):@(h)(h)(h):@(h)(h)(h)(h):@(h)(h)(h):@(h)(h)(h)(h)(h):@(h)(h)(h):@(h)(h)(h)(h):@(h)(h)(h):@(h)(h)(h)(h):@:@:@:@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)"
Case "v"
SendMessag "(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h):@(h)(h)(h)(h):@(h)(h)(h):@(h)(h)(h)(h):@(h)(h)(h):@(h)(h)(h)(h):@(h)(h)(h):@(h)(h)(h)(h)(h):@(h):@(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)"
Case "w"
SendMessag "(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h):@(h)(h)(h)(h):@(h)(h)(h):@(h)(h)(h)(h):@(h)(h)(h):@(h)(h)(h)(h):@(h):@(h):@(h)(h)(h)(h):@(h):@(h):@(h)(h)(h)(h):@(h):@(h):@(h)(h)(h)(h(h):@)(h):@(h)(h)(h)(h)(h)(h)(h(h)(h)(h)(h)(h)"
Case "x"
SendMessag "(h)(h)(h)(h)(h)(h)(h)(h)(h)(H)(h):@(h)(h)(h)(h)(h):@(h(H))(h)(h):@(h)(h)(h):@(h)(h)(H)(h)(h)(h):@(h):@(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h):@(h):@(h)(h)(h)(h)(h)(h):@(h)(h)(h):@(h)(h)(h)(h):@(h)(h)(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(H)"
Case "y"
SendMessag "(h)(h)(h)(h)(h)(h)(h)(h)(h)(H)(h):@(h)(h)(h)(h)(h):@(h)(h)(h)(h):@(h)(h)(h):@(h)(h)(h)(h)(h)(h):@(h):@(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h):@(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(h)(H)"
Case "z"
SendMessag "(h)(h)(h)(h)(h)(h)(h)(h)(h)(H)(h):@:@:@:@:@:@(h)(H)(h)(h)(h)(h)(h)(h):@(h)(h)(H)(h)(h)(h)(h)(h):@(h)(h)(h)(H)(h)(h)(h)(h):@(h)(h)(h)(h)(H)(h)(h)(h):@(h)(h)(h)(h)(h)(H)(h)(h):@(h)(h)(h)(h)(h)(h)(H)(h):@:@:@:@:@:@(h)(H)(h)(h)(h)(h)(h)(h)(h)(h)(h)(H)"
Case "007"
SendMessag "(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(D)(D)(D)(D)(*)(D)(D)(D)(D)(*)(D)(D)(D)(D)(D)(*)(*)(D)(*)(*)(D)(*)(D)(*)(*)(D)(*)(*)(*)(*)(D)(*)(*)(*)(D)(*)(*)(D)(*)(D)(*)(*)(D)(*)(*)(*)(D)(*)(*)(*)(*)(D)(*)(*)(D)(*)(D)(*)(*)(D)(*)(*)(D)(*)(*)(*)(*)(*)(D)(D)(D)(D)(*)(D)(D)(D)(D)(*)(D)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)"
Case "msn"
SendMessag "(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(8)(*)(*)(*)(8)(*)(8)(8)(8)(8)(*)(8)(8)(8)(8)(*)(*)(*)(8)(8)(*)(8)(8)(*)(8)(*)(*)(*)(*)(8)(*)(*)(8)(*)(*)(*)(8)(*)(8)(*)(8)(*)(8)(8)(8)(8)(*)(8)(*)(*)(8)(*)(*)(*)(8)(*)(*)(*)(8)(*)(*)(*)(*)(8)(*)(8)(*)(*)(8)(*)(*)(*)(8)(*)(*)(*)(8)(*)(8)(8)(8)(8)(*)(8)(*)(*)(8)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)"

End Select
Next
10
End Sub



Private Sub sk_Click()
txtMsg.Text = "(n)(n)(n)(n)(n)(n)(n)(n)(n)(n)(n)(n)(n)(6)(6)(n)(6)(6)(n)(6)(6)(n)(n)(n)(6)(6)(n)(6)(6)(n)(6)(6)(n)(n)(n)(6)(6)(n)(6)(6)(n)(6)(6)(n)(n)(n)(6)(6)(n)(6)(6)(n)(6)(6)(n)(n)(n)(n)(6)(n)(6)(6)(6)(6)(n)(n)(n)(n)(n)(6)(6)(6)(6)(6)(6)(n)(n)(n)(n)(n)(n)(n)(6)(6)(n)(n)(n)(n)(n)(n)(n)(n)(n)(6)(6)(n)(n)(n)(n)(n)(n)(n)(n)(n)(6)(6)(n)(n)(n)(n)(n)(n)(n)(n)(n)(6)(6)(n)(n)(n)(n)(n)(n)(n)(n)(n)(6)(6)(n)(n)(n)(n)(n)(n)(n)(n)(n)(n)(n)(n)(n)(n)(n)"
End Sub




Private Sub tree_Click()
txtMsg.Text = ""
txtMsg.Text = "(m)(m)(m)(m)(m)(*)(m)(m)(m)(m)(m)(m)(m)(m)(m)(*)(*)(*)(m)(m)(m)(m)(m)(m)(m)(*)(*)(*)(*)(*)(m)(m)(m)(m)(m)(m)(m)(*)(*)(*)(m)(m)(m)(m)(m)(m)(m)(*)(*)(*)(*)(*)(m)(m)(m)(m)(m)(*)(*)(*)(*)(*)(*)(*)(m)(m)(m)(m)(m)(m)(b)(b)(b)(m)(m)(m)(m)"
End Sub

Private Sub turnoverform_Click()
cust1.Text = ReadKey("HKCU\Software\TechnoBOT\Scrolling\usr1")
cust2.Text = ReadKey("HKCU\Software\TechnoBOT\Scrolling\usr2")
cust3.Text = ReadKey("HKCU\Software\TechnoBOT\Scrolling\usr3")
cust4.Text = ReadKey("HKCU\Software\TechnoBOT\Scrolling\usr4")
cust5.Text = ReadKey("HKCU\Software\TechnoBOT\Scrolling\usr5")
cust6.Text = ReadKey("HKCU\Software\TechnoBOT\Scrolling\usr6")
cust7.Text = ReadKey("HKCU\Software\TechnoBOT\Scrolling\usr7")
cust8.Text = ReadKey("HKCU\Software\TechnoBOT\Scrolling\usr8")
cust9.Text = ReadKey("HKCU\Software\TechnoBOT\Scrolling\usr9")
sf1.Visible = True
sf2.Visible = False
othr.Visible = False
turnoverform2.Visible = True
turnoverform.Visible = False
End Sub

Private Sub turnoverform2_Click()
sf2.Visible = True
sf1.Visible = False
othr.Visible = False
turnoverform2.Visible = False
turnoverform.Visible = True
End Sub

Private Sub ur1_Click()
r = ReadKey("HKCU\Software\TechnoBOT\Scrolling\usr1")
txtMsg.Text = r
End Sub

Private Sub ur2_Click()
r = ReadKey("HKCU\Software\TechnoBOT\Scrolling\usr2")
txtMsg.Text = r
End Sub

Private Sub ur3_Click()
r = ReadKey("HKCU\Software\TechnoBOT\Scrolling\usr3")
txtMsg.Text = r
End Sub

Private Sub ur4_Click()
r = ReadKey("HKCU\Software\TechnoBOT\Scrolling\usr4")
txtMsg.Text = r
End Sub

Private Sub ur5_Click()
r = ReadKey("HKCU\Software\TechnoBOT\Scrolling\usr5")
txtMsg.Text = r
End Sub

Private Sub ur6_Click()
r = ReadKey("HKCU\Software\TechnoBOT\Scrolling\usr6")
txtMsg.Text = r
End Sub

Private Sub ur7_Click()
r = ReadKey("HKCU\Software\TechnoBOT\Scrolling\usr7")
txtMsg.Text = r
End Sub

Private Sub ur8_Click()
r = ReadKey("HKCU\Software\TechnoBOT\Scrolling\usr8")
txtMsg.Text = r
End Sub

Private Sub ur9_Click()
r = ReadKey("HKCU\Software\TechnoBOT\Scrolling\usr9")
txtMsg.Text = r
End Sub

Private Sub xmark_Click()
txtMsg.Text = ""
txtMsg.Text = "(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z):|:|:|:|:|:|:|:|:|:|:|(z)(z):|:@:@:|:|:|:|:|:@:@:|(z)(z):|:|:@:@:|:|:|:@:@:|:|(z)(z):|:|:|:@:@:|:@:@:|:|:|(z)(z):|:|:|:|:@:@:@:|:|:|:|(z)(z):|:|:|:@:@:|:@:@:|:|:|(z)(z):|:|:@:@:|:|:|:@:@:|:|(z)(z):|:@:@:|:|:|:|:|:@:@:|(z)(z):|:|:|:|:|:|:|:|:|:|:|(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)(z)"
End Sub

Private Sub zshape_Click()
txtMsg.Text = ""
txtMsg.Text = "(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(m)(m)(8)(8)(e)(e)(e)(*)(*)(m)(m)(m)(8)(8)(e)(e)(*)(*)(m)(m)(m)(m)(8)(8)(e)(*)(*)(m)(m)(m)(8)(8)(e)(e)(*)(*)(m)(m)(8)(8)(e)(e)(e)(*)(*)(m)(8)(8)(e)(e)(e)(e)(*)(*)(m)(m)(8)(8)(e)(e)(e)(*)(*)(m)(m)(m)(8)(8)(e)(e)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)"
End Sub

Function SelectTxt(custom As TextBox)
    custom = LCase(custom.Text)
    custom.SelStart = Len(custom.Text) + 1
End Function

