VERSION 5.00
Begin VB.Form credits 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5715
   LinkTopic       =   "Form1"
   Picture         =   "credits.frx":0000
   ScaleHeight     =   7905
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer2 
      Interval        =   3000
      Left            =   1080
      Top             =   3000
   End
   Begin VB.Timer Timer1 
      Interval        =   2
      Left            =   3120
      Top             =   3000
   End
   Begin VB.Image Image9 
      Height          =   1290
      Left            =   960
      Picture         =   "credits.frx":9E616
      Top             =   7920
      Width           =   3720
   End
   Begin VB.Image Image12 
      Height          =   300
      Left            =   315
      Picture         =   "credits.frx":AE048
      Top             =   10920
      Width           =   5070
   End
   Begin VB.Image Image11 
      Height          =   300
      Left            =   480
      Picture         =   "credits.frx":B2FEA
      Top             =   2520
      Width           =   4680
   End
   Begin VB.Image Image10 
      Height          =   345
      Left            =   1080
      Picture         =   "credits.frx":B794C
      Top             =   10440
      Width           =   3825
   End
   Begin VB.Image Image8 
      Height          =   1035
      Left            =   2880
      Picture         =   "credits.frx":BBE8E
      Top             =   9360
      Width           =   1080
   End
   Begin VB.Image Image7 
      Height          =   1035
      Left            =   1680
      Picture         =   "credits.frx":BF908
      Top             =   9360
      Width           =   1080
   End
   Begin VB.Image Image6 
      Height          =   675
      Left            =   1200
      Picture         =   "credits.frx":C3382
      Top             =   5760
      Width           =   3450
   End
   Begin VB.Image Image4 
      Height          =   540
      Left            =   840
      Picture         =   "credits.frx":CAD68
      Top             =   1440
      Width           =   3930
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   320
      Picture         =   "credits.frx":D1C7A
      Top             =   7560
      Width           =   5070
   End
   Begin VB.Image Image3 
      Height          =   300
      Left            =   315
      Picture         =   "credits.frx":D6C1C
      Top             =   240
      Width           =   5070
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   315
      Picture         =   "credits.frx":DBBBE
      Top             =   3600
      Width           =   5070
   End
   Begin VB.Image Image5 
      Height          =   240
      Left            =   720
      Picture         =   "credits.frx":E0B60
      Top             =   5040
      Width           =   4380
   End
End
Attribute VB_Name = "credits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If
End Sub



Private Sub Image7_Click()
ShellExecute hwnd, "open", "http://www.webtech.tk", vbNullString, vbNullString, conSwNormal
End Sub

Private Sub Image8_Click()
ShellExecute hwnd, "open", "http://www.webtech.tk", vbNullString, vbNullString, conSwNormal
End Sub

Private Sub Label1_Click()

End Sub

Private Sub Timer1_Timer()
If Image9.Top = 3120 Then
Timer2.Enabled = True
Me.Hide
technoend.Show
Else
Image1.Top = Image1.Top - 5
Image2.Top = Image2.Top - 5
Image3.Top = Image3.Top - 5
Image4.Top = Image4.Top - 5
Image5.Top = Image5.Top - 5
Image6.Top = Image6.Top - 5
Image7.Top = Image7.Top - 5
Image8.Top = Image8.Top - 5
Image9.Top = Image9.Top - 5
Image10.Top = Image10.Top - 5
Image11.Top = Image11.Top - 5
Image12.Top = Image12.Top - 5
End If
End Sub

Public Function Wait(ByVal TimeToWait As Long) 'Time In seconds
    Dim EndTime As Long
    EndTime = GetTickCount + TimeToWait * 1000 '* 1000 Cause u give seconds and GetTickCount uses Milliseconds
    Do Until GetTickCount > EndTime
        DoEvents
        Loop
    End Function

