VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form F1 
   Caption         =   "Form1"
   ClientHeight    =   7050
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   8115
   LinkTopic       =   "Form1"
   Picture         =   "technobot5.frx":0000
   ScaleHeight     =   7050
   ScaleWidth      =   8115
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   480
      Top             =   6120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame st3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Nickname Trick Page 1"
      Height          =   4215
      Left            =   2880
      TabIndex        =   68
      Top             =   1680
      Visible         =   0   'False
      Width           =   4815
      Begin VB.CommandButton Command37 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Host Nickname"
         Height          =   195
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CommandButton Command38 
         BackColor       =   &H00C0C0C0&
         Caption         =   "IP Nickname"
         Height          =   195
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton Command39 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Blank Name"
         Height          =   195
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   89
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton Command40 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Set"
         Height          =   195
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   88
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton Command41 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Set As Name"
         Height          =   195
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   87
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton Command42 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Underline"
         Height          =   195
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton Command45 
         BackColor       =   &H00C0C0C0&
         Caption         =   "On"
         Height          =   195
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   85
         Top             =   3840
         Width           =   615
      End
      Begin VB.CommandButton Command46 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Off"
         Height          =   195
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   3840
         Width           =   615
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   4200
         Top             =   3240
      End
      Begin VB.TextBox underlinenick 
         BackColor       =   &H00C0C0C0&
         Height          =   405
         Left            =   120
         TabIndex        =   83
         Top             =   600
         Width           =   2295
      End
      Begin VB.CommandButton Command34 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Flash It!"
         Height          =   195
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   3600
         Width           =   975
      End
      Begin VB.TextBox flashnick1 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   240
         TabIndex        =   81
         Text            =   "Flashing"
         Top             =   2520
         Width           =   2175
      End
      Begin VB.TextBox flashnick2 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   240
         TabIndex        =   80
         Text            =   "!"
         Top             =   2760
         Width           =   2175
      End
      Begin VB.TextBox flashnick3 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   240
         TabIndex        =   79
         Text            =   "Wid"
         Top             =   3000
         Width           =   2175
      End
      Begin VB.TextBox flashnick4 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   240
         TabIndex        =   78
         Text            =   "CP BOT!"
         Top             =   3240
         Width           =   2175
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   3960
         Top             =   3240
      End
      Begin VB.CommandButton Command35 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Stop It!"
         Height          =   195
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   3600
         Width           =   975
      End
      Begin VB.CommandButton Command36 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Status On"
         Height          =   195
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   480
         Width           =   1095
      End
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   4200
         Top             =   480
      End
      Begin VB.CommandButton Command43 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Status Off"
         Height          =   195
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   720
         Width           =   1095
      End
      Begin VB.Timer Timer4 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   4440
         Top             =   3240
      End
      Begin VB.TextBox blackname 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   2640
         TabIndex        =   74
         Text            =   "Black Name Blues"
         Top             =   1200
         Width           =   2055
      End
      Begin VB.CommandButton datenick 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Date Nickname"
         Height          =   195
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox dateme 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   2640
         TabIndex        =   72
         Top             =   2400
         Width           =   2055
      End
      Begin VB.TextBox date 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   3000
         TabIndex        =   71
         Text            =   "The Date Is: "
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   2640
         TabIndex        =   70
         Text            =   "The Time Is:"
         Top             =   3480
         Width           =   975
      End
      Begin VB.TextBox time 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3720
         TabIndex        =   69
         Top             =   3480
         Width           =   1035
      End
      Begin VB.Timer Timer5 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   1200
         Top             =   3720
      End
      Begin VB.Label Label31 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Time Nickname"
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
         Left            =   2640
         TabIndex        =   102
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Label32 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Underline Nick"
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
         TabIndex        =   101
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label33 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Other Name Features"
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
         TabIndex        =   100
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label34 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Flash Nickname"
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
         TabIndex        =   99
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label Label35 
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
         Height          =   255
         Left            =   120
         TabIndex        =   98
         Top             =   2520
         Width           =   135
      End
      Begin VB.Label Label36 
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
         Height          =   255
         Left            =   120
         TabIndex        =   97
         Top             =   2760
         Width           =   135
      End
      Begin VB.Label Label37 
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
         Height          =   255
         Left            =   120
         TabIndex        =   96
         Top             =   3000
         Width           =   135
      End
      Begin VB.Label Label38 
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
         Height          =   255
         Left            =   120
         TabIndex        =   95
         Top             =   3240
         Width           =   135
      End
      Begin VB.Line Line42 
         BorderColor     =   &H00C0C0C0&
         X1              =   2520
         X2              =   2520
         Y1              =   360
         Y2              =   4080
      End
      Begin VB.Label Label39 
         BackColor       =   &H00E0E0E0&
         Caption         =   "More Nickname Features"
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
         Left            =   2520
         TabIndex        =   94
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label41 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Black Name"
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
         Left            =   2640
         TabIndex        =   93
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Date And Time"
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
         Left            =   2640
         TabIndex        =   92
         Top             =   1800
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Multiline Nick"
      Height          =   4215
      Left            =   2880
      TabIndex        =   103
      Top             =   1680
      Visible         =   0   'False
      Width           =   4815
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Send"
         Height          =   195
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   119
         Top             =   2640
         Width           =   1575
      End
      Begin VB.OptionButton op4 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   1680
         TabIndex        =   113
         Top             =   2280
         Width           =   255
      End
      Begin VB.OptionButton op3 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   1320
         TabIndex        =   112
         Top             =   2280
         Width           =   255
      End
      Begin VB.OptionButton op2 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   960
         TabIndex        =   111
         Top             =   2280
         Width           =   255
      End
      Begin VB.OptionButton op1 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   600
         TabIndex        =   110
         Top             =   2280
         Width           =   255
      End
      Begin VB.TextBox multinick4 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   240
         TabIndex        =   109
         Text            =   "cp.suddenlaunch.com"
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox multinick3 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   240
         TabIndex        =   108
         Text            =   "@"
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox multinick2 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   240
         TabIndex        =   107
         Text            =   "Corporal Punishment"
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox multinick1 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   240
         TabIndex        =   106
         Text            =   "Visit "
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label40 
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
         Height          =   255
         Left            =   120
         TabIndex        =   118
         Top             =   1680
         Width           =   135
      End
      Begin VB.Label Label29 
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
         Height          =   255
         Left            =   120
         TabIndex        =   117
         Top             =   1320
         Width           =   135
      End
      Begin VB.Label Label28 
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
         Height          =   255
         Left            =   120
         TabIndex        =   116
         Top             =   960
         Width           =   135
      End
      Begin VB.Label Label26 
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
         Height          =   255
         Left            =   120
         TabIndex        =   115
         Top             =   600
         Width           =   135
      End
      Begin VB.Label Label25 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Only     1      2      3      4"
         Height          =   255
         Left            =   120
         TabIndex        =   114
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Multiline Nick"
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
         TabIndex        =   105
         Top             =   360
         Width           =   1215
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         X1              =   0
         X2              =   5040
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.Frame index 
      BackColor       =   &H00E0E0E0&
      Height          =   2655
      Left            =   2880
      TabIndex        =   29
      Top             =   1680
      Width           =   4815
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Programmed By E-Weaver For Corporal Punishment"
         Height          =   255
         Left            =   600
         TabIndex        =   35
         Top             =   360
         Width           =   3735
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Intro"
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
         TabIndex        =   34
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   $"technobot5.frx":B897E
         Height          =   615
         Left            =   120
         TabIndex        =   33
         Top             =   1080
         Width           =   4575
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Don't Forget"
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
         TabIndex        =   32
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Visit The CP Website @:"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label30 
         BackColor       =   &H00E0E0E0&
         Caption         =   "http://www.cp.spyw.com"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   2040
         TabIndex        =   30
         Top             =   2160
         Width           =   1935
      End
   End
   Begin VB.Frame st1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Standard Features"
      Height          =   4215
      Left            =   2880
      TabIndex        =   36
      Top             =   1680
      Visible         =   0   'False
      Width           =   4815
      Begin VB.CommandButton Command17 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Contact List"
         Height          =   195
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Command18 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Allow List"
         Height          =   195
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Command19 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Reverse List"
         Height          =   195
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Command20 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Block List"
         Height          =   195
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Command21 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Status Flood"
         Height          =   195
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton Command22 
         BackColor       =   &H00C0C0C0&
         Caption         =   "On/offline Flood"
         Height          =   195
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox namechange 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   360
         TabIndex        =   53
         Text            =   "Visit "
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox namechange2 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   360
         TabIndex        =   52
         Text            =   "us"
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox namechange3 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   360
         TabIndex        =   51
         Text            =   "@"
         Top             =   2040
         Width           =   2055
      End
      Begin VB.TextBox namechange4 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   360
         TabIndex        =   50
         Text            =   "w w w .w e b t e c h . t k"
         Top             =   2400
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   720
         TabIndex        =   49
         Top             =   3000
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Option1"
         Height          =   255
         Left            =   1080
         TabIndex        =   48
         Top             =   3000
         Width           =   255
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Option1"
         Height          =   255
         Left            =   1440
         TabIndex        =   47
         Top             =   3000
         Width           =   255
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Option1"
         Height          =   255
         Left            =   1800
         TabIndex        =   46
         Top             =   3000
         Width           =   255
      End
      Begin VB.CommandButton sendfriendlybomb 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Friendly Name- off/on Flood"
         Height          =   255
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   3360
         Width           =   2175
      End
      Begin VB.CommandButton online 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Online"
         Height          =   255
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CommandButton offline 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Offline"
         Height          =   255
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   1680
         Width           =   1575
      End
      Begin VB.CommandButton logout 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Logout"
         Height          =   255
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CommandButton away 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Away"
         Height          =   255
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   2160
         Width           =   1575
      End
      Begin VB.CommandButton onthephone 
         BackColor       =   &H00C0C0C0&
         Caption         =   "On The Phone"
         Height          =   255
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   2400
         Width           =   1575
      End
      Begin VB.CommandButton brb 
         BackColor       =   &H00C0C0C0&
         Caption         =   "BRB"
         Height          =   255
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   2640
         Width           =   1575
      End
      Begin VB.CommandButton busy 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Busy"
         Height          =   255
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   2880
         Width           =   1575
      End
      Begin VB.CommandButton outtolunch 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Out To Lunch"
         Height          =   255
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "View Contact List"
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
         TabIndex        =   67
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Flodding"
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
         Left            =   120
         TabIndex        =   66
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Only     1      2      3      4"
         Height          =   255
         Left            =   240
         TabIndex        =   65
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Label Label11 
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
         Height          =   255
         Left            =   240
         TabIndex        =   64
         Top             =   1320
         Width           =   135
      End
      Begin VB.Label Label14 
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
         Height          =   255
         Left            =   240
         TabIndex        =   63
         Top             =   1680
         Width           =   135
      End
      Begin VB.Label Label15 
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
         Height          =   255
         Left            =   240
         TabIndex        =   62
         Top             =   2040
         Width           =   135
      End
      Begin VB.Label Label16 
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
         Height          =   255
         Left            =   240
         TabIndex        =   61
         Top             =   2400
         Width           =   135
      End
      Begin VB.Line Line35 
         BorderColor     =   &H00C0C0C0&
         X1              =   2640
         X2              =   2640
         Y1              =   1440
         Y2              =   3600
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Status Change"
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
         Left            =   2760
         TabIndex        =   60
         Top             =   1200
         Width           =   1455
      End
   End
   Begin VB.Frame st2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Standard Features Page 2"
      Height          =   4215
      Left            =   2880
      TabIndex        =   11
      Top             =   1680
      Visible         =   0   'False
      Width           =   4815
      Begin VB.TextBox imwc 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Text            =   "technobot5.frx":B8A0F
         Top             =   480
         Width           =   4455
      End
      Begin VB.CommandButton IMW 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Change"
         Height          =   255
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1560
         Width           =   1695
      End
      Begin VB.CommandButton offenabled 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ON"
         Height          =   255
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton offunenabled 
         BackColor       =   &H00C0C0C0&
         Caption         =   "OFF"
         Height          =   255
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton Command26 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Away Responder"
         Height          =   255
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2760
         Width           =   1335
      End
      Begin VB.CommandButton Command28 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Contact List"
         Height          =   195
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   3720
         Width           =   1455
      End
      Begin VB.CommandButton Command29 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Blocked PPL"
         Height          =   195
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   3720
         Width           =   1455
      End
      Begin VB.CommandButton Command30 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Allow List"
         Height          =   195
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   3720
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   360
         TabIndex        =   14
         Text            =   "Visit www.cp.spyw.com"
         Top             =   3360
         Width           =   4335
      End
      Begin VB.CommandButton buttonopeninbox 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Open Inbox"
         Height          =   255
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton Command33 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pick Color"
         Height          =   255
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2760
         Width           =   1095
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   4080
         Top             =   1800
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "IM Warning Changer"
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
         Left            =   360
         TabIndex        =   28
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Appear Offline And Still Talk"
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
         Left            =   360
         TabIndex        =   27
         Top             =   1920
         Width           =   2775
      End
      Begin VB.Label Label22 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Away Responder"
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
         Left            =   360
         TabIndex        =   26
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label24 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Mass Msg"
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
         Left            =   360
         TabIndex        =   25
         Top             =   3120
         Width           =   975
      End
      Begin VB.Line Line39 
         BorderColor     =   &H00C0C0C0&
         X1              =   2040
         X2              =   2040
         Y1              =   3120
         Y2              =   2520
      End
      Begin VB.Label Label23 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Open Inbox"
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
         TabIndex        =   24
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Line Line40 
         BorderColor     =   &H00C0C0C0&
         X1              =   3360
         X2              =   3360
         Y1              =   3120
         Y2              =   2520
      End
      Begin VB.Label Label27 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Color Changer"
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
         Left            =   3480
         TabIndex        =   23
         Top             =   2520
         Width           =   1335
      End
   End
   Begin VB.Image Image7 
      Height          =   180
      Left            =   7200
      Picture         =   "technobot5.frx":B8A2B
      Top             =   100
      Width           =   180
   End
   Begin VB.Image Image6 
      Height          =   180
      Left            =   7440
      Picture         =   "technobot5.frx":B8C1D
      Top             =   105
      Width           =   180
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "  "
      Height          =   135
      Left            =   7680
      TabIndex        =   121
      Top             =   120
      Width           =   135
   End
   Begin VB.Image Image2 
      Height          =   1170
      Left            =   2880
      Picture         =   "technobot5.frx":B8E0F
      Top             =   480
      Width           =   4605
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Other"
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
      Left            =   480
      TabIndex        =   10
      Top             =   3280
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Standard Features"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label changeme 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome To Techno BOT!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   6720
      Width           =   7095
   End
   Begin VB.Label bar 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Emotion Flooder"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   435
      TabIndex        =   7
      Top             =   4755
      Width           =   1965
   End
   Begin VB.Label bar5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "E-Mail Bomber"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   435
      TabIndex        =   5
      Top             =   4275
      Width           =   1965
   End
   Begin VB.Label bar1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   465
      TabIndex        =   4
      Top             =   1470
      Width           =   1935
   End
   Begin VB.Label bar3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   470
      TabIndex        =   2
      Top             =   2595
      Width           =   1930
   End
   Begin VB.Image Image3 
      Height          =   465
      Left            =   360
      Picture         =   "technobot5.frx":CA7D9
      Top             =   4200
      Width           =   2115
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "........Techno BOT Created By Dreamingweb"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label menu 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Menu"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   480
      Width           =   855
   End
   Begin VB.Label bar2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   465
      TabIndex        =   3
      Top             =   2040
      Width           =   1930
   End
   Begin VB.Label bar4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Emotion Machine"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   435
      TabIndex        =   6
      Top             =   3795
      Width           =   1965
   End
   Begin VB.Image Image1 
      Height          =   465
      Left            =   360
      Picture         =   "technobot5.frx":CDB73
      Top             =   3720
      Width           =   2115
   End
   Begin VB.Image bar6 
      Height          =   465
      Left            =   360
      Picture         =   "technobot5.frx":D0F0D
      Top             =   4680
      Width           =   2115
   End
   Begin VB.Label xbar 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Patches"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   435
      TabIndex        =   104
      Top             =   5235
      Width           =   1965
   End
   Begin VB.Image Image4 
      Height          =   465
      Left            =   360
      Picture         =   "technobot5.frx":D42A7
      Top             =   5160
      Width           =   2115
   End
   Begin VB.Label creditbar 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Credits"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   435
      TabIndex        =   120
      Top             =   5715
      Width           =   1965
   End
   Begin VB.Image Image5 
      Height          =   465
      Left            =   360
      Picture         =   "technobot5.frx":D7641
      Top             =   5640
      Width           =   2115
   End
End
Attribute VB_Name = "F1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents Msn As MsgrObject
Attribute Msn.VB_VarHelpID = -1
Dim aUser As IMsgrUser
Dim Header As String


Private Sub Command17_Click()
contactlst.Show
End Sub

Private Sub Command18_Click()
contactlst.Show
End Sub

Private Sub Command19_Click()
contactlst.Show
End Sub





Private Sub Command20_Click()
contactlst.Show
End Sub

Private Sub Command3_Click()
If op1.Value = True Then
Msn.Services.PrimaryService.FriendlyName = "­" & multinick1.Text
End If
If op2.Value = True Then
Msn.Services.PrimaryService.FriendlyName = "­" & multinick1.Text & "­" & multinick2.Text
End If
If op3.Value = True Then
Msn.Services.PrimaryService.FriendlyName = "­" & multinick1.Text & "­" & multinick2.Text & "­" & multinick3.Text
End If
If op4.Value = True Then
Msn.Services.PrimaryService.FriendlyName = "­" & multinick1.Text & "­" & multinick2.Text & "­" & multinick3.Text & "­" & multinick4.Text
End If
End Sub

Private Sub Command34_Click()
Timer2.Enabled = True
changeme.Caption = "Flashing Friendly Name Set"
End Sub

Private Sub Command35_Click()
Timer2.Enabled = False
Timer4.Enabled = False
changeme.Caption = "Flashing Friendly Name Stopped"
End Sub

Private Sub Command36_Click()
Timer3.Enabled = True
Wait 1
changeme.Caption = "Friendly Name: " & Msn.Services.PrimaryService.FriendlyName 'check this too
End Sub

Private Sub Command37_Click()
Msn.Services.PrimaryService.FriendlyName = Winsock1.LocalHostName
changeme.Caption = "Host Friendly Name Set: " & Winsock1.LocalHostName
End Sub

Private Sub Command38_Click()
Msn.Services.PrimaryService.FriendlyName = Winsock1.LocalIP
changeme.Caption = "IP Friendly Name Set: " & Winsock1.LocalIP
End Sub

Private Sub Command39_Click()
Msn.Services.PrimaryService.FriendlyName = " "
changeme.Caption = "Blank Friendly Name Set"
End Sub

Private Sub Command40_Click()
Msn.Services.PrimaryService.FriendlyName = Chr(10) & Chr(13) & Chr(147) & Chr(10) & " " & blackname.Text & " " & Chr(147)
changeme.Caption = "Black Friendly Name Set"
End Sub

Private Sub Command41_Click()
changeme.Caption = "Friendly Name Changed To: " & underlinenick.FontUnderline 'might be wrong
Msn.Services.PrimaryService.FriendlyName = underlinenick.Text
End Sub

Private Sub Command42_Click()
underlinenick.Text = underlinenick.Text & "¯¯¯¯¯¯¯¯¯¯¯¯"
End Sub

Private Sub Command43_Click()
Timer3.Enabled = False
changeme.Caption = "Friendly Name Status Off"
End Sub



Private Sub datenick_Click()
xyz = date & dateme
Msn.Services.PrimaryService.FriendlyName = xyz
changeme.Caption = "Date Friendly Name Set: " & xyz
End Sub

Private Sub Command45_Click()
Timer5.Enabled = True
changeme.Caption = "Time Friendly Name Set"
End Sub

Private Sub Command46_Click()
Timer5.Enabled = False
changeme.Caption = "Time Friendly Name Stopped"
End Sub

Private Sub Form_Load()
If Me.Picture <> 0 Then
  Call SetAutoRgn(Me)
End If
dateme.Text = Format(Now, "dddd, mmmm dd, yyyy")
offunenabled.Enabled = False
op1.Value = True
Set Msn = New MsgrObject
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbLeftButton Then
 ReleaseCapture
  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If
End Sub
'button settings
Private Sub bar1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim l As ColorConstants
l = &H80000011
p = &HC0C0C0
bar1.BackColor = l
bar.BackColor = p
bar2.BackColor = p
bar3.BackColor = p
bar4.BackColor = p
bar5.BackColor = p
xbar.BackColor = p
creditbar.BackColor = p
changeme.Caption = "First lot of standard features"
st2.Visible = True
st1.Visible = False
st3.Visible = False
index.Visible = False
Frame1.Visible = False
End Sub
Private Sub bar2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim l As ColorConstants
q = &H80000011
p = &HC0C0C0
bar2.BackColor = q
bar1.BackColor = p
bar.BackColor = p
bar3.BackColor = p
bar4.BackColor = p
bar5.BackColor = p
xbar.BackColor = p
creditbar.BackColor = p
changeme.Caption = "Second load of standard features"
st2.Visible = False
st1.Visible = True
st3.Visible = False
index.Visible = False
Frame1.Visible = False
End Sub
Private Sub bar3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim w As ColorConstants
w = &H80000011
p = &HC0C0C0
bar2.BackColor = p
bar1.BackColor = p
bar.BackColor = p
bar4.BackColor = p
bar5.BackColor = p
bar3.BackColor = w
xbar.BackColor = p
creditbar.BackColor = p
changeme.Caption = "Third Load Of Standard Features"
st2.Visible = False
st1.Visible = False
st3.Visible = True
index.Visible = False
Frame1.Visible = False
End Sub
Private Sub bar4_Click()
picgen.Show
End Sub
Private Sub bar5_Click()
Form1.Show
End Sub
Private Sub bar4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim e As ColorConstants
e = &H80000011
p = &HC0C0C0
bar4.BackColor = e
bar3.BackColor = p
bar2.BackColor = p
bar1.BackColor = p
bar.BackColor = p
bar5.BackColor = p
xbar.BackColor = p
creditbar.BackColor = p
changeme.Caption = "Use the emotion machine to generate big emotion pictures"
End Sub
Private Sub bar5_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim f As ColorConstants
f = &H80000011
p = &HC0C0C0
bar5.BackColor = f
bar4.BackColor = p
bar3.BackColor = p
bar2.BackColor = p
bar1.BackColor = p
bar.BackColor = p
xbar.BackColor = p
creditbar.BackColor = p
changeme.Caption = "E-Mail bomb those lamers wid this"
End Sub
Private Sub bar_Click()
scroll.Show
End Sub
Private Sub bar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim u As ColorConstants
u = &H80000011
p = &HC0C0C0
bar.BackColor = u
bar5.BackColor = p
bar4.BackColor = p
bar3.BackColor = p
bar2.BackColor = p
bar1.BackColor = p
xbar.BackColor = p
changeme.Caption = "Flood with big emotion pictures!!!"
End Sub





Private Sub xbar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim u As ColorConstants
u = &H80000011
p = &HC0C0C0
bar.BackColor = p
bar5.BackColor = p
bar4.BackColor = p
bar3.BackColor = p
bar2.BackColor = p
bar1.BackColor = p
xbar.BackColor = u
creditbar.BackColor = p
changeme.Caption = "Add patches to your messenger- Get the new smiley emotion pictures!"
st3.Visible = False
st2.Visible = False
index.Visible = True
End Sub
Private Sub creditbar_click()
credits.Show
End Sub
Private Sub creditbar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim u As ColorConstants
u = &H80000011
p = &HC0C0C0
bar.BackColor = p
bar5.BackColor = p
bar4.BackColor = p
bar3.BackColor = p
bar2.BackColor = p
bar1.BackColor = p
xbar.BackColor = p
creditbar.BackColor = u
changeme.Caption = "Credits.....Includes my contact details and thanks!"
st3.Visible = False
st2.Visible = False
index.Visible = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim l As ColorConstants
l = &HC0C0C0
q = &HC0C0C0
e = &HC0C0C0
w = &HC0C0C0
f = &HC0C0C0
u = &HC0C0C0
bar1.BackColor = l
bar2.BackColor = q
bar3.BackColor = w
bar4.BackColor = e
bar5.BackColor = f
bar.BackColor = u
xbar.BackColor = u
creditbar.BackColor = l
changeme.Caption = "Dreamingwebs Technobot"
End Sub



Private Sub imw_Click()
    CreateKey "HKLM\Software\Microsoft\MessengerService\Policies\IMWarning", imwc.Text 'This tells the location to write the info, text1 = what u want it to say
changeme.Caption = "IM Warning Changed"
End Sub



Private Sub Label20_Click()
End
End Sub

Private Sub offenabled_Click()
On Error Resume Next 'error crap
    For Each User In Msn.List(MLIST_ALLOW)
    Msn.List(MLIST_ALLOW).Remove User 'appear offline
    Next
    offenabled.Enabled = False 'disables the enable button
    offunenabled.Enabled = True 'enables the disable button
    changeme.Caption = "You Are Currently Shown As Offline & Still Able To Chat"
End Sub



Private Sub offunenabled_Click()
 On Error Resume Next 'error crap
    For Each User In Msn.List(MLIST_REVERSE)
    Msn.List(MLIST_ALLOW).Add User 'Appear online
    Next
    Command2.Enabled = False 'disables the disable button
    Command1.Enabled = True 'enables the enable button
    changeme.Caption = "You Are Now Shown As Online"
End Sub
Private Sub buttonopeninbox_Click()
MessengerAPI.Messenger.OpenInbox
changeme.Caption = "Inbox Opened"
End Sub
Private Sub Command33_Click()
changeme.Caption = "Currently Changing Chat Font Color"
CommonDialog1.ShowColor
regWriteSubKey HKEY_CURRENT_USER, "Software\Microsoft\MessengerService", "IM Color", CommonDialog1.Color
changeme.Caption = "Chat Font Color Changed Successfully"
End Sub
Private Sub Command28_Click()
changeme.Caption = "Mass Message Of Contact List In Process..."
On Error Resume Next
   MassMsg = Text1.Text
   For Each User In Msn.List(MLIST_CONTACT)
User.SendText "MIME-Version: 1.0" & vbCrLf & "Content-Type: text/plain; charset=UTF-8" & vbCrLf & "X-MMS-IM-Format: EF=; CO=0000FF; CS=0; PF=12" & vbCrLf & vbCrLf, MassMsg, MMSGTYPE_NORESULT
changeme.Caption = "Mass Message Of Contact List Completed"
Next
changeme.Caption = "Mass Message Of Contacts List Unsuccessful"
End Sub
Private Sub Command29_Click()
changeme.Caption = "Mass Message Of Blocked Contacts List In Process..."
On Error Resume Next
   MassMsg = Text1.Text
   For Each User In Msn.List(MLIST_BLOCK)
User.SendText "MIME-Version: 1.0" & vbCrLf & "Content-Type: text/plain; charset=UTF-8" & vbCrLf & "X-MMS-IM-Format: EF=; CO=0000FF; CS=0; PF=12" & vbCrLf & vbCrLf, MassMsg, MMSGTYPE_NORESULT
changeme.Caption = "Mass Message Of Blocked Contacts List Completed"
Next
changeme.Caption = "Mass Message Of Blocked Contacts List Unsuccessful"
End Sub
Private Sub Command30_Click()
changeme.Caption = "Mass Message Of Allowed Chat Contacts In Process..."
On Error Resume Next
   MassMsg = Text1.Text
   For Each User In Msn.List(MLIST_ALLOW)
User.SendText "MIME-Version: 1.0" & vbCrLf & "Content-Type: text/plain; charset=UTF-8" & vbCrLf & "X-MMS-IM-Format: EF=; CO=0000FF; CS=0; PF=12" & vbCrLf & vbCrLf, MassMsg, MMSGTYPE_NORESULT
changeme.Caption = "Mass Message Of Allowed Chat Contacts In Completed"
Next
changeme.Caption = "Mass Message Of Allowed Chat Contacts In Unsuccessful"
End Sub
Private Sub Command21_Click()
changeme.Caption = "Techno Bot Status Flodding In Progress"
Msn.LocalState = MSTATE_AWAY

Msn.LocalState = MSTATE_BUSY
Msn.LocalState = MSTATE_ON_THE_PHONE
Msn.LocalState = MSTATE_OUT_TO_LUNCH
Msn.LocalState = MSTATE_BE_RIGHT_BACK
Msn.LocalState = MSTATE_ONLINE
Msn.LocalState = MSTATE_AWAY
Msn.LocalState = MSTATE_BUSY
Msn.LocalState = MSTATE_ON_THE_PHONE
Msn.LocalState = MSTATE_OUT_TO_LUNCH
Msn.LocalState = MSTATE_BE_RIGHT_BACK
Msn.LocalState = MSTATE_ONLINE
changeme.Caption = "Techno Bot Status Flood Completed"
End Sub
Private Sub Command22_Click()
changeme.Caption = "Techno Bot On/Offline Flood In Progress"
Msn.LocalState = MSTATE_ONLINE
Msn.LocalState = MSTATE_INVISIBLE
Msn.LocalState = MSTATE_ONLINE
Msn.LocalState = MSTATE_INVISIBLE
Msn.LocalState = MSTATE_ONLINE
Msn.LocalState = MSTATE_INVISIBLE
Msn.LocalState = MSTATE_ONLINE
Msn.LocalState = MSTATE_INVISIBLE
Msn.LocalState = MSTATE_ONLINE
Msn.LocalState = MSTATE_INVISIBLE
Msn.LocalState = MSTATE_ONLINE
Msn.LocalState = MSTATE_INVISIBLE
changeme.Caption = "Techno Bot On/Offline Flood Completed"
End Sub

Private Sub sendfriendlybomb_Click()
changeme.Caption = "Friendly Name Bomb Started"
If Option1.Value = True Then
OnlyFirst
Hold 1
End If
If Option2.Value = True Then
OnlySecond
Hold 1
End If
If Option3.Value = True Then
OnlyThird
Hold 1
End If
If Option4.Value = True Then
OnlyFourth
Hold 1
End If
changeme.Caption = "Friendly Name Bomb Stopped"
End Sub
Sub OnlyFirst()
changeme.Caption = "One Friendly Name Flood In Progress"
Msn.LocalState = MSTATE_ONLINE
Msn.LocalState = MSTATE_INVISIBLE
Msn.LocalState = MSTATE_ONLINE
Msn.Services.PrimaryService.FriendlyName = namechange.Text
changeme.Caption = "One Friendly Name Flood Completed"
End Sub
Sub OnlySecond()
changeme.Caption = "Two Friendly Name Floods In Progress"
Msn.LocalState = MSTATE_ONLINE
Msn.LocalState = MSTATE_INVISIBLE
Msn.LocalState = MSTATE_ONLINE
Msn.Services.PrimaryService.FriendlyName = namechange2.Text
Msn.LocalState = MSTATE_INVISIBLE
Msn.LocalState = MSTATE_ONLINE
Msn.Services.PrimaryService.FriendlyName = namechange.Text
changeme.Caption = "Two Friendly Name Floods Completed"
End Sub
Sub OnlyThird()
changeme.Caption = "Third Friendly Name Flood In Progress"
Msn.LocalState = MSTATE_ONLINE
Msn.LocalState = MSTATE_INVISIBLE
Msn.LocalState = MSTATE_ONLINE
Msn.Services.PrimaryService.FriendlyName = namechange3.Text
Msn.LocalState = MSTATE_INVISIBLE
Msn.LocalState = MSTATE_ONLINE
Msn.Services.PrimaryService.FriendlyName = namechange2.Text
Msn.LocalState = MSTATE_INVISIBLE
Msn.LocalState = MSTATE_ONLINE
Msn.Services.PrimaryService.FriendlyName = namechange.Text
changeme.Caption = "Three Friendly Name Floods Completed"
End Sub
Sub OnlyFourth()
changeme.Caption = "Four Friendly Name Floods In Progress"
Msn.LocalState = MSTATE_ONLINE
Msn.LocalState = MSTATE_INVISIBLE
Msn.LocalState = MSTATE_ONLINE
Msn.Services.PrimaryService.FriendlyName = namechange4.Text
Msn.LocalState = MSTATE_INVISIBLE
Msn.LocalState = MSTATE_ONLINE
Msn.Services.PrimaryService.FriendlyName = namechange3.Text
Msn.LocalState = MSTATE_INVISIBLE
Msn.LocalState = MSTATE_ONLINE
Msn.Services.PrimaryService.FriendlyName = namechange2.Text
Msn.LocalState = MSTATE_INVISIBLE
Msn.LocalState = MSTATE_ONLINE
Msn.Services.PrimaryService.FriendlyName = namechange.Text
changeme.Caption = "Four Friendly Name Floods Completed"
End Sub
Private Sub online_Click()
Msn.LocalState = MSTATE_ONLINE
changeme.Caption = "Your Status Is Now Online"
End Sub
Private Sub onthephone_Click()
Msn.LocalState = MSTATE_ON_THE_PHONE
changeme.Caption = "Your Status Is Now On The Phone"
End Sub
Private Sub outtolunch_Click()
Msn.LocalState = MSTATE_OUT_TO_LUNCH
changeme.Caption = "Your Status Is Now Out To Lunch"
End Sub
Private Sub logout_Click()
MessengerAPI.Messenger.Signout
changeme.Caption = "Your Are Logged Out Of Windows Messenger"
End Sub
Private Sub away_Click()
Msn.LocalState = MSTATE_AWAY
changeme.Caption = "Your Status Is Now Away"
End Sub
Private Sub brb_Click()
Msn.LocalState = MSTATE_BE_RIGHT_BACK
changeme.Caption = "Your Status Is Now Be Right Back"
End Sub

Private Sub busy_Click()
Msn.LocalState = MSTATE_BUSY
changeme.Caption = "Your Status Is Now Busy"
End Sub
Private Sub offline_Click()
Msn.LocalState = MSTATE_INVISIBLE
changeme.Caption = "Your Status Is Now Invisible (Offline)"
End Sub

Private Sub Timer1_Timer()
time.Text = Format(Now, "hh:mm:ss AM/PM")
End Sub

Private Sub Timer2_Timer()
Msn.Services.PrimaryService.FriendlyName = flashnick1.Text
        Wait 1
Msn.Services.PrimaryService.FriendlyName = flashnick2.Text
        Wait 1
Msn.Services.PrimaryService.FriendlyName = flashnick3.Text
        Wait 1
Msn.Services.PrimaryService.FriendlyName = flashnick4.Text
End Sub

Private Sub Timer3_Timer()
If Msn.LocalState = MSTATE_ONLINE Then     ' If Online
Msn.Services.PrimaryService.FriendlyName = "I'm Online!"
End If
If Msn.LocalState = MSTATE_AWAY Then      ' If Away
Msn.Services.PrimaryService.FriendlyName = "Away"
End If
If Msn.LocalState = MSTATE_BUSY Then         ' If Busy
Msn.Services.PrimaryService.FriendlyName = "Busy"
End If
If Msn.LocalState = MSTATE_BE_RIGHT_BACK Then    ' If BRB
Msn.Services.PrimaryService.FriendlyName = "Be Right Back"
End If
If Msn.LocalState = MSTATE_ON_THE_PHONE Then     ' If OTP
Msn.Services.PrimaryService.FriendlyName = "On The Phone"
End If
If Msn.LocalState = MSTATE_OUT_TO_LUNCH Then
Msn.Services.PrimaryService.FriendlyName = "Out To Lunch"
End If
If Msn.LocalState = MSTATE_INVISIBLE Then
Msn.Services.PrimaryService.FriendlyName = "Appearing Offline"
End If
End Sub

Private Sub Timer5_Timer()
time.Text = Format(Now, "hh:mm:ss AM/PM")
Msn.Services.PrimaryService.FriendlyName = time.Text
End Sub

Public Function Wait(ByVal TimeToWait As Long) 'Time In seconds
    Dim EndTime As Long
    EndTime = GetTickCount + TimeToWait * 1000 '* 1000 Cause u give seconds and GetTickCount uses Milliseconds
    Do Until GetTickCount > EndTime
        DoEvents
        Loop
    End Function
    Public Function Hold(ByVal TimeToWait As Long) 'Time In seconds
    Dim EndTime As Long
    EndTime = GetTickCount + TimeToWait * 500 '* 1000 Cause u give seconds and GetTickCount uses Milliseconds
    Do Until GetTickCount > EndTime
        DoEvents
        Loop
    End Function


