VERSION 5.00
Begin VB.Form findnreplace 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Find And Replace"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6675
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   6675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton FindButton 
      Caption         =   "Find"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3840
      TabIndex        =   8
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton FindNextButton 
      Caption         =   "Find Again"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5040
      TabIndex        =   7
      Top             =   375
      Width           =   1215
   End
   Begin VB.CommandButton ReplaceButton 
      Caption         =   "Replace"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3840
      TabIndex        =   6
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton ReplaceAllButton 
      Caption         =   "Replace All"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5040
      TabIndex        =   5
      Top             =   855
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   5040
      TabIndex        =   4
      Top             =   1575
      Width           =   1230
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Case sensitive"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   330
      TabIndex        =   3
      Top             =   1530
      Width           =   2040
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Whole word only"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   330
      TabIndex        =   2
      Top             =   1830
      Width           =   2040
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   375
      Width           =   2115
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   855
      Width           =   2100
   End
   Begin VB.Label Label1 
      Caption         =   "Find what"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   270
      TabIndex        =   10
      Top             =   375
      Width           =   1410
   End
   Begin VB.Label Label2 
      Caption         =   "Replace with"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   855
      Width           =   1485
   End
End
Attribute VB_Name = "findnreplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Position As Integer

Private Sub FindButton_Click()
Dim FindFlags As Integer

    Position = 0
    FindFlags = Check1.Value * 4 + Check2.Value * 2
    Position = picgen.textcode.Find(Text1.Text, Position + 1, , FindFlags)
    If Position >= 0 Then
        ReplaceButton.Enabled = True
        ReplaceAllButton.Enabled = True
        picgen.SetFocus
    Else
        MsgBox "String not found", vbOKOnly, "Emotion Machine Error"
        ReplaceButton.Enabled = False
        ReplaceAllButton.Enabled = False
    End If
End Sub

Private Sub FindNextButton_Click()
Dim FindFlags

FindFlags = Check1.Value * 4 + Check2.Value * 2
Position = picgen.textcode.Find(Text1.Text, Position + 1, , FindFlags)
If Position > 0 Then
    picgen.SetFocus
Else
    MsgBox "String not found", vbOKOnly, "Search Help"
    ReplaceButton.Enabled = False
    ReplaceAllButton.Enabled = False
End If

End Sub

Private Sub Command5_Click()

    findnreplace.Hide
    
End Sub
Private Sub Form_GotFocus()
Text1.SetFocus
End Sub
Private Sub ReplaceButton_Click()
Dim FindFlags As Integer

    picgen.textcode.SelText = Text2.Text
    FindFlags = Check1.Value * 4 + Check2.Value * 2
    Position = picgen.textcode.Find(Text1.Text, Position + 1, , FindFlags)
    If Position > 0 Then
        picgen.SetFocus
    Else
        MsgBox "String not found", vbOKOnly, "Search Help"
        ReplaceButton.Enabled = False
        ReplaceAllButton.Enabled = False
    End If
    
End Sub

Private Sub ReplaceAllButton_Click()
Dim FindFlags As Integer

    FindFlags = Check1.Value * 4 + Check2.Value * 2
    picgen.textcode.SelText = Text2.Text
    Position = picgen.textcode.Find(Text1.Text, Position + 1, , FindFlags)
    While Position > 0
        picgen.textcode.SelText = Text2.Text
        Position = picgen.textcode.Find(Text1.Text, Position + 1, , FindFlags)
    Wend
        ReplaceButton.Enabled = False
        ReplaceAllButton.Enabled = False
        MsgBox "Done replacing", vbOKOnly, "Search Help"
End Sub


