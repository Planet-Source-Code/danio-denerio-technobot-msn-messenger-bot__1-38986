VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form picgen 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Technobot Emotion Machine Build 5"
   ClientHeight    =   7095
   ClientLeft      =   150
   ClientTop       =   420
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "widundo4.frx":0000
   ScaleHeight     =   7095
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton getreal 
      Caption         =   "Get Real Code"
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   5520
      Width           =   2655
   End
   Begin VB.CommandButton find 
      Caption         =   "Find"
      Height          =   255
      Left            =   5280
      TabIndex        =   14
      Top             =   6120
      Width           =   2415
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   240
      Top             =   240
   End
   Begin VB.CommandButton open 
      Caption         =   "Open"
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   6120
      Width           =   2415
   End
   Begin VB.CommandButton new 
      Caption         =   "New"
      Height          =   255
      Left            =   2880
      TabIndex        =   10
      Top             =   6120
      Width           =   2295
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton clearbox 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Clear Box"
      Height          =   255
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5520
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Emotion Tool Box"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5175
      Left            =   5880
      TabIndex        =   5
      Top             =   600
      Width           =   1815
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Space"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   360
         TabIndex        =   9
         Top             =   4800
         Width           =   1140
      End
      Begin VB.Image happy 
         Height          =   285
         Left            =   240
         Picture         =   "widundo4.frx":BA2AE
         Top             =   2400
         Width           =   285
      End
      Begin VB.Image beer 
         Height          =   285
         Left            =   1320
         Picture         =   "widundo4.frx":BA764
         Top             =   2760
         Width           =   285
      End
      Begin VB.Image thumbdown 
         Height          =   285
         Left            =   960
         Picture         =   "widundo4.frx":BAC1A
         Top             =   2400
         Width           =   285
      End
      Begin VB.Image female 
         Height          =   285
         Left            =   600
         Picture         =   "widundo4.frx":BB0D0
         Top             =   2760
         Width           =   285
      End
      Begin VB.Image wink 
         Height          =   285
         Left            =   960
         Picture         =   "widundo4.frx":BB586
         Top             =   2040
         Width           =   285
      End
      Begin VB.Image male 
         Height          =   285
         Left            =   960
         Picture         =   "widundo4.frx":BBA3C
         Top             =   2760
         Width           =   285
      End
      Begin VB.Image unhappy 
         Height          =   285
         Left            =   600
         Picture         =   "widundo4.frx":BBEF2
         Top             =   2040
         Width           =   285
      End
      Begin VB.Image drink 
         Height          =   285
         Left            =   240
         Picture         =   "widundo4.frx":BC3A8
         Top             =   2760
         Width           =   285
      End
      Begin VB.Image thumbup 
         Height          =   285
         Left            =   600
         Picture         =   "widundo4.frx":BC85E
         Top             =   2400
         Width           =   285
      End
      Begin VB.Image tounge 
         Height          =   285
         Left            =   1320
         Picture         =   "widundo4.frx":BCD14
         Top             =   2400
         Width           =   285
      End
      Begin VB.Image oh 
         Height          =   285
         Left            =   1320
         Picture         =   "widundo4.frx":BD1CA
         Top             =   2040
         Width           =   285
      End
      Begin VB.Image malehands 
         Height          =   285
         Left            =   960
         Picture         =   "widundo4.frx":BD680
         Top             =   3120
         Width           =   285
      End
      Begin VB.Image veryhappy 
         Height          =   285
         Left            =   240
         Picture         =   "widundo4.frx":BDB36
         Top             =   3120
         Width           =   285
      End
      Begin VB.Image bat 
         Height          =   285
         Left            =   1320
         Picture         =   "widundo4.frx":BDFEC
         Top             =   3120
         Width           =   285
      End
      Begin VB.Image femalehands 
         Height          =   285
         Left            =   600
         Picture         =   "widundo4.frx":BE4A2
         Top             =   3120
         Width           =   285
      End
      Begin VB.Image straghtface 
         Height          =   285
         Left            =   240
         Picture         =   "widundo4.frx":BE958
         Top             =   1680
         Width           =   285
      End
      Begin VB.Image upset 
         Height          =   285
         Left            =   240
         Picture         =   "widundo4.frx":BEE0E
         Top             =   2040
         Width           =   285
      End
      Begin VB.Image coffee 
         Height          =   285
         Left            =   600
         Picture         =   "widundo4.frx":BF2C4
         Top             =   3480
         Width           =   285
      End
      Begin VB.Image mooon 
         Height          =   285
         Left            =   1320
         Picture         =   "widundo4.frx":BF77A
         Top             =   3840
         Width           =   285
      End
      Begin VB.Image cry 
         Height          =   285
         Left            =   600
         Picture         =   "widundo4.frx":BFC30
         Top             =   1680
         Width           =   285
      End
      Begin VB.Image blush 
         Height          =   285
         Left            =   960
         Picture         =   "widundo4.frx":C00E6
         Top             =   1680
         Width           =   285
      End
      Begin VB.Image cat 
         Height          =   285
         Left            =   1320
         Picture         =   "widundo4.frx":C059C
         Top             =   3480
         Width           =   285
      End
      Begin VB.Image Gboy 
         Height          =   285
         Left            =   960
         Picture         =   "widundo4.frx":C0A52
         Top             =   1320
         Width           =   285
      End
      Begin VB.Image dog 
         Height          =   285
         Left            =   960
         Picture         =   "widundo4.frx":C0F08
         Top             =   3480
         Width           =   285
      End
      Begin VB.Image star 
         Height          =   285
         Left            =   960
         Picture         =   "widundo4.frx":C13BE
         Top             =   3840
         Width           =   285
      End
      Begin VB.Image angry 
         Height          =   285
         Left            =   1320
         Picture         =   "widundo4.frx":C1874
         Top             =   1320
         Width           =   285
      End
      Begin VB.Image glasses 
         Height          =   285
         Left            =   1320
         Picture         =   "widundo4.frx":C1D2A
         Top             =   1680
         Width           =   285
      End
      Begin VB.Image phone 
         Height          =   285
         Left            =   240
         Picture         =   "widundo4.frx":C21E0
         Top             =   600
         Width           =   285
      End
      Begin VB.Image film 
         Height          =   285
         Left            =   600
         Picture         =   "widundo4.frx":C2696
         Top             =   600
         Width           =   285
      End
      Begin VB.Image flower 
         Height          =   285
         Left            =   240
         Picture         =   "widundo4.frx":C2B4C
         Top             =   960
         Width           =   285
      End
      Begin VB.Image photo 
         Height          =   285
         Left            =   960
         Picture         =   "widundo4.frx":C3002
         Top             =   600
         Width           =   285
      End
      Begin VB.Image dropflower 
         Height          =   285
         Left            =   1320
         Picture         =   "widundo4.frx":C34B8
         Top             =   600
         Width           =   285
      End
      Begin VB.Image brokenheart 
         Height          =   285
         Left            =   1320
         Picture         =   "widundo4.frx":C396E
         Top             =   960
         Width           =   285
      End
      Begin VB.Image present 
         Height          =   285
         Left            =   600
         Picture         =   "widundo4.frx":C3E24
         Top             =   960
         Width           =   285
      End
      Begin VB.Image devil 
         Height          =   285
         Left            =   600
         Picture         =   "widundo4.frx":C42DA
         Top             =   1320
         Width           =   285
      End
      Begin VB.Image heart 
         Height          =   285
         Left            =   240
         Picture         =   "widundo4.frx":C4790
         Top             =   1320
         Width           =   285
      End
      Begin VB.Image lips 
         Height          =   285
         Left            =   960
         Picture         =   "widundo4.frx":C4C46
         Top             =   960
         Width           =   285
      End
      Begin VB.Image idea 
         Height          =   285
         Left            =   240
         Picture         =   "widundo4.frx":C50FC
         Top             =   3480
         Width           =   285
      End
      Begin VB.Image cake 
         Height          =   285
         Left            =   960
         Picture         =   "widundo4.frx":C55B2
         Top             =   4200
         Width           =   285
      End
      Begin VB.Image email 
         Height          =   285
         Left            =   240
         Picture         =   "widundo4.frx":C5A68
         Top             =   3840
         Width           =   285
      End
      Begin VB.Image note 
         Height          =   285
         Left            =   600
         Picture         =   "widundo4.frx":C5F1E
         Top             =   3840
         Width           =   285
      End
      Begin VB.Image time 
         Height          =   285
         Left            =   600
         Picture         =   "widundo4.frx":C63D4
         Top             =   4200
         Width           =   285
      End
      Begin VB.Image mess 
         Height          =   285
         Left            =   240
         Picture         =   "widundo4.frx":C688A
         Top             =   4200
         Width           =   285
      End
      Begin VB.Image rainbow 
         Height          =   255
         Left            =   1320
         Picture         =   "widundo4.frx":C6D40
         Top             =   4200
         Width           =   330
      End
      Begin VB.Image asl 
         Height          =   330
         Left            =   600
         Picture         =   "widundo4.frx":C7206
         Top             =   240
         Width           =   300
      End
      Begin VB.Image cuffs 
         Height          =   330
         Left            =   960
         Picture         =   "widundo4.frx":C7770
         Top             =   240
         Width           =   315
      End
      Begin VB.Image sun 
         Height          =   285
         Left            =   1320
         Picture         =   "widundo4.frx":C7D32
         Top             =   240
         Width           =   285
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "New Line"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   360
         TabIndex        =   6
         Top             =   4560
         Width           =   1140
      End
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Send To E-Weaver"
      Height          =   255
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6480
      Width           =   2415
   End
   Begin VB.CommandButton codewindow 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Code Window"
      Height          =   255
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5160
      Width           =   2655
   End
   Begin VB.CommandButton graphicwindow 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Graphic Window"
      Height          =   255
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5160
      Width           =   2655
   End
   Begin RichTextLib.RichTextBox textcode 
      Height          =   4335
      Left            =   360
      TabIndex        =   15
      Top             =   600
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   7646
      _Version        =   393217
      TextRTF         =   $"widundo4.frx":C81E8
   End
   Begin RichTextLib.RichTextBox act 
      Height          =   4335
      Left            =   360
      TabIndex        =   16
      Top             =   600
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   7646
      _Version        =   393217
      ReadOnly        =   -1  'True
      TextRTF         =   $"widundo4.frx":C826A
   End
   Begin RichTextLib.RichTextBox rt 
      Height          =   4335
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   7646
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"widundo4.frx":C82EC
   End
   Begin VB.CommandButton saveasrcode 
      Caption         =   "SaveAS (RCode)"
      Height          =   255
      Left            =   2880
      TabIndex        =   19
      Top             =   6480
      Width           =   2295
   End
   Begin VB.CommandButton saveas 
      Caption         =   "SaveAS (Code)"
      Height          =   255
      Left            =   2880
      TabIndex        =   12
      Top             =   6480
      Width           =   2295
   End
   Begin VB.CommandButton savercode 
      Caption         =   "Save RCode"
      Height          =   255
      Left            =   360
      TabIndex        =   18
      Top             =   6480
      Width           =   2415
   End
   Begin VB.CommandButton save 
      Caption         =   "Save Code"
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   6480
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   285
      Left            =   7200
      Picture         =   "widundo4.frx":C836E
      Top             =   5400
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Saving"
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
      Left            =   360
      TabIndex        =   7
      Top             =   5880
      Width           =   615
   End
   Begin VB.Label emotionmachine 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Graphic Window Viewer"
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
      Left            =   360
      TabIndex        =   0
      Top             =   330
      Width           =   7215
   End
End
Attribute VB_Name = "picgen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub angry_Click() ' only one wrong
textcode.Text = textcode.Text & ":@=" 'change
GetClipboard
Clipboard.SetData angry.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub

Private Sub asl_Click()
textcode.Text = textcode.Text & "(?)="
GetClipboard
Clipboard.SetData asl.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub

Private Sub bat_Click()
textcode.Text = textcode.Text & ":[="
GetClipboard
Clipboard.SetData bat.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub

Private Sub beer_Click()
textcode.Text = textcode.Text & "(b)="
GetClipboard
Clipboard.SetData beer.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub

Private Sub blush_Click()
textcode.Text = textcode.Text & ":$="
GetClipboard
Clipboard.SetData blush.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub

Private Sub brokenheart_Click()
textcode.Text = textcode.Text & "(u)=" 'change
GetClipboard
Clipboard.SetData brokenheart.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub

Private Sub cake_Click()
textcode.Text = textcode.Text & "(^)=" 'change
GetClipboard
Clipboard.SetData cake.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub

Private Sub cat_Click()
textcode.Text = textcode.Text & "(@)=" 'change
GetClipboard
Clipboard.SetData cat.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub

Private Sub coffee_Click()
textcode.Text = textcode.Text & "(c)=" 'change
GetClipboard
Clipboard.SetData coffee.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub





Private Sub graphicwindow_Click()
emotionmachine.Caption = "Graphic Window Viewer"
savercode.Visible = False
saveasrcode.Visible = False
save.Visible = True
saveas.Visible = True
Bitch
rt.Visible = True
textcode.Visible = False
act.Visible = False
End Sub

Private Sub codewindow_Click()
emotionmachine.Caption = "Code Window"
savercode.Visible = False
saveasrcode.Visible = False
save.Visible = True
saveas.Visible = True
textcode.Visible = True
rt.Visible = False
act.Visible = False
End Sub
Function SetClipboard()
'Clipboard.SetData = clip2
'Clipboard.SetText = clip
End Function
Function GetClipboard()
Clipboard.Clear
'clip = Clipboard.GetText
'clip2 = Clipboard.GetData
End Function

Private Sub Command3_Click()
getcode
Bitch
If act.Text = "" Then
MsgBox "There Are No Emotions Added To Create A Emotion Picture " & vbCrLf & "Please Retry!", vbExclamation, "Emotion Picture Machine"
GoTo Opps
End If
emaildream.Show
Opps:
End Sub

Private Sub clearbox_Click()
rt.TextRTF = ""
textcode.Text = ""
act.Text = ""
End Sub

Private Sub getreal_Click()
emotionmachine.Caption = "Windows Messenger (Real) Code Window"
savercode.Visible = True
saveasrcode.Visible = True
save.Visible = False
saveas.Visible = False
act.Text = ""
Bitch
getcode
act.Visible = True
rt.Visible = False
textcode.Visible = False
End Sub



Private Sub cry_Click()
textcode.Text = textcode.Text & ":'(=" 'change
GetClipboard
Clipboard.SetData cry.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub

Private Sub cuffs_Click()
textcode.Text = textcode.Text & "(%)=" 'change
GetClipboard
Clipboard.SetData cuffs.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub

Private Sub devil_Click()
textcode.Text = textcode.Text & "(6)=" 'change
GetClipboard
Clipboard.SetData devil.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub

Private Sub dog_Click()
textcode.Text = textcode.Text & "(&)=" 'change
GetClipboard
Clipboard.SetData dog.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub

Private Sub drink_Click()
textcode.Text = textcode.Text & "(d)=" 'change
GetClipboard
Clipboard.SetData drink.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub

Private Sub dropflower_Click()
textcode.Text = textcode.Text & "(w)=" 'change
GetClipboard
Clipboard.SetData dropflower.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub

Private Sub email_Click()
textcode.Text = textcode.Text & "(e)=" 'change
GetClipboard
Clipboard.SetData email.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub

Private Sub female_Click()
textcode.Text = textcode.Text & "(x)=" 'change
GetClipboard
Clipboard.SetData female.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub

Private Sub femalehands_Click()
textcode.Text = textcode.Text & "(})=" 'change
GetClipboard
Clipboard.SetData femalehands.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub

Private Sub film_Click()
textcode.Text = textcode.Text & "(~)=" 'change
GetClipboard
Clipboard.SetData film.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub

Private Sub find_Click()
findnreplace.Show
End Sub

Private Sub flower_Click()
textcode.Text = textcode.Text & "(f)=" 'change
GetClipboard
Clipboard.SetData flower.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub

Private Sub rt_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
rt.Locked = True
End Sub
Private Sub frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
rt.Locked = False
End Sub
Private Sub Gboy_Click()
textcode.Text = textcode.Text & "(a)=" 'change
GetClipboard
Clipboard.SetData Gboy.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub

Private Sub glasses_Click()
textcode.Text = textcode.Text & "(h)=" 'change
GetClipboard
Clipboard.SetData glasses.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub

Private Sub happy_Click()
textcode.Text = textcode.Text & ":)=" 'change
GetClipboard
Clipboard.SetData happy.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub

Private Sub heart_Click()
textcode.Text = textcode.Text & "(l)=" 'change
GetClipboard
Clipboard.SetData heart.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub

Private Sub idea_Click()
textcode.Text = textcode.Text & "(i)=" 'change
GetClipboard
Clipboard.SetData idea.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub
Private Sub Label3_Click()
textcode.Text = textcode.Text & "nl="
rt.SelText = vbCrLf
End Sub

Private Sub Label4_Click()
textcode.Text = textcode.Text & "(q)=" ' probelly buggy check it out
GetClipboard
Clipboard.SetData Image1.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub

Private Sub lips_Click()
textcode.Text = textcode.Text & "(k)=" 'change
GetClipboard
Clipboard.SetData lips.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub

Private Sub male_Click()
textcode.Text = textcode.Text & "(z)=" 'change
GetClipboard
Clipboard.SetData male.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub

Private Sub malehands_Click()
textcode.Text = textcode.Text & "({)=" 'change
GetClipboard
Clipboard.SetData malehands.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub

Private Sub mess_Click()
textcode.Text = textcode.Text & "(m)=" 'change
GetClipboard
Clipboard.SetData mess.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub

Private Sub mooon_Click()
textcode.Text = textcode.Text & "(S)="
GetClipboard
Clipboard.SetData mooon.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub

Private Sub note_Click()
textcode.Text = textcode.Text & "(8)=" 'change
GetClipboard
Clipboard.SetData Note.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub

Private Sub oh_Click()
textcode.Text = textcode.Text & ":o=" 'change
GetClipboard
Clipboard.SetData oh.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub

Private Sub phone_Click()
textcode.Text = textcode.Text & "(t)=" 'change
GetClipboard
Clipboard.SetData phone.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub

Private Sub photo_Click()
textcode.Text = textcode.Text & "(p)=" 'change
GetClipboard
Clipboard.SetData photo.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub

Private Sub present_Click()
textcode.Text = textcode.Text & "(g)=" 'change
GetClipboard
Clipboard.SetData present.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub

Private Sub rainbow_Click()
textcode.Text = textcode.Text & "(r)=" 'change
GetClipboard
Clipboard.SetData rainbow.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub

Private Sub saveasrcode_Click()
On Error GoTo 15
'--------------------
'CommonDialog1.ShowSave Renders The Save DialogBox
CommonDialog1.ShowSave
'The Open Command Loads A File Into RAM[ For Output As (#1(FileMarker While Open)).You Can Use #(AnyNumber) Consistantly Through The Sub
Open CommonDialog1.FileName For Output As #1
'--------------------
'You Will Be Extracting File With The Open Command Later, Remember The Order You Save It In
'Write To File
'Example To Write More: Write#1, Text1.Text, Text2.Text, Button1.Caption    etc......
Write #1, act.Text
'--------------------
'It Is Important To Close Open Files At The End of The Sub, Otherwise Errors Will Occur
Close #1
'-------------------
'Now That We Have Opened A File We Can Make The Save Button Visible
save.Enabled = True
'--------------------
15
End Sub

Private Sub savercode_Click()
'On Error Goto Command For Cancel In The Dialog Box
On Error GoTo 14
'-------------------
'The Open Command Loads A File Into RAM[ For Output As (#1(FileMarker While Open)).You Can Use #(AnyNumber) Consistantly Through The Sub
Open CommonDialog1.FileName For Output As #1
'--------------------
'You Will Be Extracting File With The Open Command Later, Remember The Order You Save It In
'Write To File
'Example To Write More: Write#1, Text1.Text, Text2.Text, Button1.Caption    etc......
Write #1, act.Text

'--------------------
'It Is Important To Close Open Files At The End of The Sub, Otherwise Errors Will Occur
Close #1
'--------------------
' If An Error Occurs, The Computer Will Resume Processing The Code For This Command At Line 10
14
End Sub

Private Sub star_Click()
textcode.Text = textcode.Text & "(*)=" 'change
GetClipboard
Clipboard.SetData star.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub

Private Sub straghtface_Click()
textcode.Text = textcode.Text & ":|=" 'change
GetClipboard
Clipboard.SetData straghtface.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub

Private Sub sun_Click()
textcode.Text = textcode.Text & "(#)=" 'change
GetClipboard
Clipboard.SetData sun.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub
Private Sub thumbdown_Click()
textcode.Text = textcode.Text & "(y)=" 'change
GetClipboard
Clipboard.SetData thumbdown.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub

Private Sub thumbup_Click()
textcode.Text = textcode.Text & "(n)=" 'change
GetClipboard
Clipboard.SetData thumbup.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub

Private Sub time_Click()
textcode.Text = textcode.Text & "(o)=" 'change
GetClipboard
Clipboard.SetData time.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub

Private Sub Timer1_Timer()
If textcode.Visible = True Then
find.Enabled = True
End If
If textcode.Visible = False Then
find.Enabled = False
End If
End Sub

Private Sub tounge_Click()
textcode.Text = textcode.Text & ":p=" 'change
GetClipboard
Clipboard.SetData tounge.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub


Private Sub unhappy_Click()
textcode.Text = textcode.Text & ":(=" 'change
GetClipboard
Clipboard.SetData unhappy.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub

Private Sub upset_Click()
textcode.Text = textcode.Text & ":s=" 'change
GetClipboard
Clipboard.SetData upset.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub

Private Sub veryhappy_Click()
textcode.Text = textcode.Text & ":d=" 'change
GetClipboard
Clipboard.SetData veryhappy.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub

Private Sub wink_Click()
textcode.Text = textcode.Text & ";)=" 'change
GetClipboard
Clipboard.SetData wink.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
End Sub
'for saving stuff
Private Sub rt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
Sendkey vbKeyDown
End If
If KeyCode = vbKeyLeft Then
Sendkey vbKeyRight
End If
End Sub
Private Sub form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
Sendkey vbKeyDown
End If
If KeyCode = vbKeyLeft Then
Sendkey vbKeyRight
End If
End Sub
Private Sub Form_Load()
act.Visible = False
rt.Visible = True
textcode.Visible = False
savercode.Visible = False
saveasrcode.Visible = False
find.Enabled = False
'If rt.asl.
'This Will Become Visible After You Open Or Save As..
'The CommonDialog1.FileName Property Should Be Blank At This Point, Which is Why You Cant Click Save
save.Enabled = False
savercode.Enabled = False
'--------------------
'Set A Filter For The Dialog box (Available File Types To Save As/Open)
CommonDialog1.Filter = "Text Files|*.TXT"
'--------------------
End Sub


Private Sub Form_Resize()
'On Error Goto Command For Minimize & Not Visible Properties Of Form
On Error GoTo 10
'Adjust The Height Of The RichTextBox, As You Resize The Form
rt.Height = Examples.Height - StatusBar1.Height - 750
'Adjust The Width Of The RichTextBox As You Resize The Form
rt.Width = Examples.Width - 100
'If You Minimize The Form The Computer Will Skip To Line 10
10
End Sub

Private Sub new_Click()
'Clear The RichTextBox
textcode.Text = ""
rt.TextRTF = ""
act.Visible = False
rt.Visible = True
textcode.Visible = False
'------------------
'Disable Save Function
save.Enabled = False
End Sub

Private Sub open_Click()
On Error GoTo 10
'On Error Goto Command For Cancel In The Dialog Box
'On Error GoTo 10
'CommonDialog1.ShowOpen Renders The Open DialogBox
CommonDialog1.ShowOpen
'-------------------
'CommonDialog1.FileName Is The Value Of The Path & Url The User Selected
'The Open Command Loads A File Into RAM[ For Input As (#1(FileMarker While Open)).You Can Use #(AnyNumber) Consistantly Through The Sub
Open CommonDialog1.FileName For Input As #1
'-------------------
'The Input Statement Extracts Information From A File From The Same Order You Saved It In
'It Is Important to Extract That Information In The Same Order You Saved It In
'To input More Inforemation: Ex: Input #1, VariableA, VariableB,VariableC
'Text1.Text = VariableA: Text2.Text = VariableB: Text3.Text = Button1.Caption   etc...
Input #1, AnyVariableWillDo$
textcode.Text = ""
'-------------------
'Apply The Extracted Information To The TextBox
textcode.Text = AnyVariableWillDo$
'-------------------
'Refresh The RichTextBox
textcode.Refresh
act.Visible = False
rt.Visible = True
textcode.Visible = False
'-------------------
'It Is Important To Close Open Files At The End of The Sub, Otherwise Errors Will Occur
Close #1
'-------------------
'Now That We Have Opened A File We Can Make The Save Button Visible
save.Enabled = True
'Apply The File Name To The StatusBar
Bitch
10
End Sub
Sub Bitch()
Dim spliter As Variant
spliter = Split(textcode.Text, "=")
rt.Locked = False
rt.Text = ""
For i = 0 To UBound(spliter) - 1
rt.SelStart = Len(rt.Text)
Select Case spliter(i)
Case "nl"
rt.SelText = rt.SelText & vbCrLf
Case "(q)"
GetClipboard
Clipboard.SetData Image1.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
Case "(?)"
GetClipboard
Clipboard.SetData asl.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
Case "(%)"
GetClipboard
Clipboard.SetData cuffs.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
Case "(#)"
GetClipboard
Clipboard.SetData sun.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
Case "(t)"
GetClipboard
Clipboard.SetData phone.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
Case ":@"
GetClipboard
Clipboard.SetData angry.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
Case "(?)"
GetClipboard
Clipboard.SetData asl.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
Case ":["
GetClipboard
Clipboard.SetData bat.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
Case "(b)"
GetClipboard
Clipboard.SetData beer.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
Case ":$"
GetClipboard
Clipboard.SetData blush.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
Case "(u)"
GetClipboard
Clipboard.SetData brokenheart.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
Case "(^)"
GetClipboard
Clipboard.SetData cake.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
Case "(@)"
GetClipboard
Clipboard.SetData cat.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
Case "(c)"
GetClipboard
Clipboard.SetData coffee.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
Case ":'("
GetClipboard
Clipboard.SetData cry.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard



Case "(6)"
GetClipboard
Clipboard.SetData devil.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard


Case "(&)"
GetClipboard
Clipboard.SetData dog.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard

Case "(d)"
GetClipboard
Clipboard.SetData drink.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard


Case "(w)"
GetClipboard
Clipboard.SetData dropflower.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard

Case "(e)"
GetClipboard
Clipboard.SetData email.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard

Case "(x)"
GetClipboard
Clipboard.SetData female.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard


Case "(})"
GetClipboard
Clipboard.SetData femalehands.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard


Case "(~)"
GetClipboard
Clipboard.SetData film.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard

Case "(f)"
GetClipboard
Clipboard.SetData flower.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard

Case "(a)"
GetClipboard
Clipboard.SetData Gboy.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard

Case "(h)"
GetClipboard
Clipboard.SetData glasses.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard

Case ":)"
GetClipboard
Clipboard.SetData happy.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard

Case "(l)"
GetClipboard
Clipboard.SetData heart.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard

Case "(i)"
GetClipboard
Clipboard.SetData idea.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard

Case "(k)"
GetClipboard
Clipboard.SetData lips.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard

Case "(z)"
GetClipboard
Clipboard.SetData male.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard

Case "({)"
GetClipboard
Clipboard.SetData malehands.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard

Case "(m)"
GetClipboard
Clipboard.SetData mess.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard

Case "(S)"
GetClipboard
Clipboard.SetData mooon.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard

Case "(8)"
GetClipboard
Clipboard.SetData Note.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard

Case ":o"
GetClipboard
Clipboard.SetData oh.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard

Case "(t)"
GetClipboard
Clipboard.SetData phone.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard

Case "(p)"
GetClipboard
Clipboard.SetData photo.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard

Case "(g)"
GetClipboard
Clipboard.SetData present.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard

Case "(r)"
GetClipboard
Clipboard.SetData rainbow.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard

Case "(*)"
GetClipboard
Clipboard.SetData star.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard

Case ":|"
GetClipboard
Clipboard.SetData straghtface.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard

Case "(#)"
GetClipboard
Clipboard.SetData sun.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard

Case "(n)"
GetClipboard
Clipboard.SetData thumbdown.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard

Case "(y)"
GetClipboard
Clipboard.SetData thumbup.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard

Case "(o)"
GetClipboard
Clipboard.SetData time.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard

Case ":p"
GetClipboard
Clipboard.SetData tounge.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard

Case ":("
GetClipboard
Clipboard.SetData unhappy.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard

Case ":s"
GetClipboard
Clipboard.SetData upset.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard

Case ":d"
GetClipboard
Clipboard.SetData veryhappy.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
Case "."
rt.TextRTF = rt.TextRTF & "try"
Case ";)"
GetClipboard
Clipboard.SetData wink.Picture
SendMessage rt.hwnd, WM_PASTE, 0, 0
SetClipboard
rt.Locked = True
End Select
Next
10
End Sub
Sub getcode()
Dim spliter As Variant
spliter = Split(textcode.Text, "=")
rt.Text = ""
For i = 0 To UBound(spliter) - 1
rt.SelStart = Len(rt.Text)
Select Case spliter(i)
Case "(q)"
act.SelText = act.SelText & "   "
Case "(?)"
act.SelText = act.SelText & "(?)"
Case "(%)"
act.SelText = act.SelText & "(%)"
Case "(#)"
act.SelText = act.SelText & "(#)"
Case "(t)"
act.SelText = act.SelText & "(t)"
Case ":@"
act.SelText = act.SelText & ":@"
Case ":["
act.SelText = act.SelText & ":["
Case "(b)"
act.SelText = act.SelText & "(b)"
Case ":$"
act.SelText = act.SelText & ":$"
Case "(u)"
act.SelText = act.SelText & "(u)"
Case "(^)"
act.SelText = act.SelText & "(^)"
Case "(@)"
act.SelText = act.SelText & "(@)"
Case "(c)"
act.SelText = act.SelText & "(c)"
Case ":'("
act.SelText = act.SelText & ":'("
Case "(6)"
act.SelText = act.SelText & "(6)"
Case "(&)"
act.SelText = act.SelText & "(&)"
Case "(d)"
act.SelText = act.SelText & "(d)"
Case "(w)"
act.SelText = act.SelText & "(w)"
Case "(e)"
act.SelText = act.SelText & "(e)"
Case "(x)"
act.SelText = act.SelText & "(x)"
Case "(})"
act.SelText = act.SelText & "(})"
Case "(~)"
act.SelText = act.SelText & "(~)"
Case "(f)"
act.SelText = act.SelText & "(f)"
Case "(a)"
act.SelText = act.SelText & "(a)"
Case "(h)"
act.SelText = act.SelText & "(h)"
Case ":)"
act.SelText = act.SelText & ":)"
Case "(l)"
act.SelText = act.SelText & "(l)"
Case "(i)"
act.SelText = act.SelText & "(i)"
Case "(k)"
act.SelText = act.SelText & "(k)"
Case "(z)"
act.SelText = act.SelText & "(z)"
Case "({)"
act.SelText = act.SelText & "({)"
Case "(m)"
act.SelText = act.SelText & "(m)"
Case "(S)"
act.SelText = act.SelText & "(S)"
Case "(8)"
act.SelText = act.SelText & "(8)"
Case ":o"
act.SelText = act.SelText & ":o"
Case "(t)"
act.SelText = act.SelText & "(t)"
Case "(p)"
act.SelText = act.SelText & "(p)"
Case "(g)"
act.SelText = act.SelText & "(g)"
Case "(r)"
act.SelText = act.SelText & "(r)"
Case "(*)"
act.SelText = act.SelText & "(*)"
Case ":|"
act.SelText = act.SelText & ":|"
Case "(#)"
act.SelText = act.SelText & "(#)"
Case "(n)"
act.SelText = act.SelText & "(n)"
Case "(y)"
act.SelText = act.SelText & "(y)"
Case "(o)"
act.SelText = act.SelText & "(o)"
Case ":p"
act.SelText = act.SelText & ":p"
Case ":("
act.SelText = act.SelText & ":("
Case ":s"
act.SelText = act.SelText & ":s"
Case ":d"
act.SelText = act.SelText & ":d"
Case "."
rt.TextRTF = rt.TextRTF & "try"
Case ";)"
act.SelText = act.SelText & ";)"
Case "nl"
act.SelText = act.SelText & vbCrLf
Case "vbCrLf"
act.SelText = act.SelText & vbCrLf
End Select
Next
10
End Sub
Private Sub save_Click()
'On Error Goto Command For Cancel In The Dialog Box
On Error GoTo 10
'-------------------
'The Open Command Loads A File Into RAM[ For Output As (#1(FileMarker While Open)).You Can Use #(AnyNumber) Consistantly Through The Sub
Open CommonDialog1.FileName For Output As #1
'--------------------
'You Will Be Extracting File With The Open Command Later, Remember The Order You Save It In
'Write To File
'Example To Write More: Write#1, Text1.Text, Text2.Text, Button1.Caption    etc......
Write #1, textcode.Text
'--------------------
'It Is Important To Close Open Files At The End of The Sub, Otherwise Errors Will Occur
Close #1
'--------------------
' If An Error Occurs, The Computer Will Resume Processing The Code For This Command At Line 10
10
End Sub

Private Sub saveas_Click()
'On Error Goto Command For Cancel In The Dialog Box
On Error GoTo 11
'--------------------
'CommonDialog1.ShowSave Renders The Save DialogBox
CommonDialog1.ShowSave
'The Open Command Loads A File Into RAM[ For Output As (#1(FileMarker While Open)).You Can Use #(AnyNumber) Consistantly Through The Sub
Open CommonDialog1.FileName For Output As #1
'--------------------
'You Will Be Extracting File With The Open Command Later, Remember The Order You Save It In
'Write To File
'Example To Write More: Write#1, Text1.Text, Text2.Text, Button1.Caption    etc......
Write #1, textcode.Text
'--------------------
'It Is Important To Close Open Files At The End of The Sub, Otherwise Errors Will Occur
Close #1
'-------------------
'Now That We Have Opened A File We Can Make The Save Button Visible
save.Enabled = True
'--------------------
11
End Sub

'for macro

