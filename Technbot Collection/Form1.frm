VERSION 5.00
Begin VB.Form contactlst 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form1"
   ClientHeight    =   6405
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4080
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   6405
   ScaleWidth      =   4080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton bl 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Block List"
      Height          =   495
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5520
      Width           =   855
   End
   Begin VB.CommandButton al 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Allow List"
      Height          =   495
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5520
      Width           =   855
   End
   Begin VB.CommandButton cl 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Contacts List"
      Height          =   495
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5520
      Width           =   855
   End
   Begin VB.CommandButton rl 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Reverse List"
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5520
      Width           =   855
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00E0E0E0&
      Height          =   4740
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   2895
   End
   Begin VB.Label title 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
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
      Left            =   600
      TabIndex        =   1
      Top             =   480
      Width           =   2895
   End
End
Attribute VB_Name = "contactlst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents messengerobj As MsgrObject
Attribute messengerobj.VB_VarHelpID = -1
Public MsgUsers As IMsgrUsers

Private Sub al_Click()
List1.Clear
title.Caption = "Your Allow List"
GetPeople List1
End Sub
Private Sub bl_Click()
List1.Clear
title.Caption = "Your Blocked List"
MeBlocked List1
End Sub

Private Sub cl_Click()
List1.Clear
title.Caption = "Your Contacts List"
MyContact List1
End Sub

Private Sub Form_Load()
Set messengerobj = New MsgrObject
If Me.Picture <> 0 Then
  Call SetAutoRgn(Me)
  End If
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If
End Sub
Public Sub GetPeople(LV As ListBox)
Set MsgUsers = messengerobj.List(MLIST_ALLOW)
For X = 0 To MsgUsers.count - 1
LV.AddItem MsgUsers.Item(X).FriendlyName
Set MsgUsers = messengerobj.List(MLIST_ALLOW)
Next X
End Sub

Public Sub GetAddedMe(LV As ListBox)
Set MsgUsers = messengerobj.List(MLIST_REVERSE)
For X = 0 To MsgUsers.count - 1
LV.AddItem MsgUsers.Item(X).FriendlyName
Set MsgUsers = messengerobj.List(MLIST_REVERSE)
Next X
End Sub
Public Sub MeBlocked(LV As ListBox)
Set MsgUsers = messengerobj.List(MLIST_BLOCK)
For X = 0 To MsgUsers.count - 1
LV.AddItem MsgUsers.Item(X).FriendlyName
Set MsgUsers = messengerobj.List(MLIST_BLOCK)
Next X
End Sub
Public Sub MyContact(LV As ListBox)
Set MsgUsers = messengerobj.List(MLIST_CONTACT)
For X = 0 To MsgUsers.count - 1
LV.AddItem MsgUsers.Item(X).FriendlyName
Set MsgUsers = messengerobj.List(MLIST_CONTACT)
Next X
End Sub
Private Sub rl_Click()
List1.Clear
title.Caption = "Reversed List"
GetAddedMe List1
End Sub
