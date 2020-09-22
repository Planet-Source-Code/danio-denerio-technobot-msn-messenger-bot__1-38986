VERSION 5.00
Begin VB.Form responder 
   Caption         =   "Auto Responder"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   6885
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Away Responder"
      Height          =   4215
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   5055
      Begin VB.CommandButton buttonaway 
         Caption         =   "&Away On"
         Height          =   375
         Left            =   1440
         TabIndex        =   5
         Top             =   3000
         Width           =   1695
      End
      Begin VB.ListBox List1 
         Height          =   840
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   4695
      End
      Begin VB.TextBox AwayMess 
         Height          =   615
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   2280
         Width           =   4695
      End
      Begin VB.OptionButton Option7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Option7"
         Height          =   255
         Left            =   480
         TabIndex        =   2
         Top             =   3240
         Width           =   255
      End
      Begin VB.OptionButton Option8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Option7"
         Height          =   255
         Left            =   3960
         TabIndex        =   1
         Top             =   3240
         Width           =   255
      End
      Begin VB.Line Line36 
         BorderColor     =   &H8000000A&
         X1              =   0
         X2              =   5040
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label Label25 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Log:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Your away message"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   1425
      End
      Begin VB.Label Label28 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Echo Text"
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
         TabIndex        =   7
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label29 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Use Away Msg"
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
         Left            =   3360
         TabIndex        =   6
         Top             =   3000
         Width           =   1455
      End
   End
End
Attribute VB_Name = "responder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public X As Integer 'for auto responder
Public AllUsers As String 'for auto responder
Public Ses As IMsgrIMSession 'for auto responder
Public blnAway As Boolean 'for auto responder
Public WithEvents Msn As MsgrObject
Attribute Msn.VB_VarHelpID = -1
Dim Header As String


Private Sub buttonaway_Click() 'for auto responder
List1.Clear

'Let's see the current mode

Select Case buttonaway.Caption

Case "&Away On" 'for auto responder
'We put away mode on by setting the boolean to True
blnAway = True
buttonaway.Caption = "&Away Off"

Case "&Away Off"
'We put away mode off by setting the boolean to False
blnAway = False
Command1.Caption = "&Away On"
End Select
End Sub


Private Sub Msg_OnTextReceived(ByVal pIMSession As Messenger.IMsgrIMSession, ByVal pSourceUser As Messenger.IMsgrUser, ByVal bstrMsgHeader As String, ByVal bstrMsgText As String, pfEnableDefault As Boolean) 'for auto responder


'Warning: You DONT see your away message yourself!


'We have received a signal that a user is typing or sending a message to us


'Set a Session variable
Set Ses = pIMSession


'make me away
If blnAway = True Then
If Option7.Value = True Then
    Ses.SendText bstrMsgHeader, bstrMsgText, MMSGTYPE_NO_RESULT
    If Option8.Value = True Then
    Ses.SendText bstrMsgHeader, AwayMess.Text, MMSGTYPE_ALL_RESULTS
    End If
List1.AddItem "[" & time & "] Away message sent to " & pSourceUser.FriendlyName
End If
End If
End Sub

'=====WARNING====='
' Ses.SendText bstrMsgHeader, AwayMess.Text, MMSGTYPE_ALL_RESULTS
' The message HEADER is the message HEADER THAT WE HAVE RECEIVED
' So if we get a message in Away mode from a dude called 'Dude' :-) ,
' we are using his headers to send a message back, so he will see in his screen 'Dude is typing a message'...
'=====WARNING====='

'About the message headers
'The headers that are used in Messenger are pretty the same as e-mail headers.
'Just enable the piece of code below:
'MsgBox bstrmsgbheader
'
'The information that is placed in the header is Font,Fontsize and more

