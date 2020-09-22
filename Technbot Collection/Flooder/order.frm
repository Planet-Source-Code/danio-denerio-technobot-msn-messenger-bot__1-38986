VERSION 5.00
Begin VB.Form order 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Order Flood Information"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   3120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "To Use Custom Big Emotion Pictures Use The Button Numbers Seperated Again By A Coma (,)"
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "To Use Default Big Emotion Pictures Enter The Button Caption Followed By A Coma (,)"
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   2895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Codes That Can Be Used Are Alphabet Codes; So If You Wished To Flood ABC You Would Enter: A,B,C,"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "To Generate A Flood Using Different Picture Types You Must Use The Codes Below Seperated By A Coma After Each Code"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "order"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
