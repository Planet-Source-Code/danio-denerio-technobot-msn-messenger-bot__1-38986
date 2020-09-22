VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form emaildream 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Send Your BIG Emotion Picture"
   ClientHeight    =   3825
   ClientLeft      =   60
   ClientTop       =   225
   ClientWidth     =   5130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSubject 
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Text            =   "**BIG EMOTION PICTURE**"
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Your Address:"
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
      Height          =   975
      Left            =   240
      TabIndex        =   8
      Top             =   0
      Width           =   4725
      Begin VB.TextBox txtSenderName 
         Height          =   285
         Left            =   1635
         TabIndex        =   0
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox txtSender 
         Height          =   285
         Left            =   1635
         TabIndex        =   1
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Name:"
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
         Left            =   120
         TabIndex        =   10
         Top             =   285
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "E-mail address:"
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
         Left            =   120
         TabIndex        =   9
         Top             =   645
         Width           =   1305
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   240
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send message"
      Height          =   375
      Left            =   840
      TabIndex        =   6
      Top             =   3000
      Width           =   3855
   End
   Begin VB.TextBox txtMessage 
      Height          =   1215
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1680
      Width           =   4695
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Reply To:"
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
      Height          =   975
      Left            =   240
      TabIndex        =   11
      Top             =   0
      Width           =   4725
      Begin VB.TextBox txtReplyToName 
         Height          =   285
         Left            =   1635
         TabIndex        =   2
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox txtReplyTo 
         Height          =   285
         Left            =   1635
         TabIndex        =   3
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Name:"
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "E-mail address:"
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   1305
      End
   End
   Begin VB.Label status 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Status: "
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
      Left            =   240
      TabIndex        =   15
      Top             =   3480
      Width           =   4695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Additional Comments:"
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
      Left            =   240
      TabIndex        =   14
      Top             =   1440
      Width           =   1830
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
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
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   1080
      Width           =   720
   End
End
Attribute VB_Name = "emaildream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This Part send the E-Mail                                                           '
'                                                                                    '
'                                                                                    '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Enum SMTP_State
    MAIL_CONNECT
    MAIL_HELO
    MAIL_FROM
    MAIL_RCPTTO
    MAIL_DATA
    MAIL_DOT
    MAIL_QUIT
End Enum

Private m_State As SMTP_State
Private m_strEncodedFiles As Variant
'
Private Sub cmdSend_Click() '
On Error GoTo Lamer
    Dim i As Integer
    Dim strServer As String, ColonPos As Integer, lngPort As Long
    '
    
    'ShowStatus

     strServer = Trim("mail.hotmail.com")
    'find out if the sender is using a Proxy server
    ColonPos = InStr(strServer, ":")
    If ColonPos = 0 Then
        'no proxy so use standard SMTP port
        Winsock1.Connect strServer, 25
    Else
        'Proxy, so get proxy port number and parse out the server name or IP address
        lngPort = CLng(Right$(strServer, Len(strServer) - ColonPos))
        strServer = Left$(strServer, ColonPos - 1)
        Winsock1.Connect strServer, lngPort
    End If
    m_State = MAIL_CONNECT
    '
Lamer:
End Sub
Private Sub Form_Load()
Frame4.Visible = False
txtMessage.Text = "Hey," & vbCrLf & "E-Weaver, I'm sending my big emotion picture to you, you may put it on your site or programs as long as i get some form of credit if you make use of my work of art" & vbCrLf & "Cheers," & vbCrLf & "__________"
End Sub



Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim strServerResponse   As String
    Dim strResponseCode     As String
    Dim strDataToSend       As String
    '
    'Retrive data from winsock buffer
    '
    Winsock1.GetData strServerResponse
    '
    Debug.Print strServerResponse
    '
    'Get server response code (first three symbols)
    '
    strResponseCode = Left(strServerResponse, 3)
    '
    'Only these three codes tell us that previous
    'command accepted successfully and we can go on
    '
    If strResponseCode = "250" Or _
       strResponseCode = "220" Or _
       strResponseCode = "354" Then
       
        Select Case m_State
            Case MAIL_CONNECT
                'Change current state of the session
                m_State = MAIL_HELO
                '
                'Remove blank spaces
                strDataToSend = Trim$(txtSender)
                '
                'Retrieve mailbox name from e-mail address
                strDataToSend = Left$(strDataToSend, _
                                InStr(1, strDataToSend, "@") - 1)
                'Send HELO command to the server
                Winsock1.SendData "HELO " & strDataToSend & vbCrLf
                '
                Debug.Print "HELO " & strDataToSend
                '
            Case MAIL_HELO
                '
                'Change current state of the session
                m_State = MAIL_FROM
                '
                'Send MAIL FROM command to the server
                Winsock1.SendData "MAIL FROM:" & Trim$(txtSender) & vbCrLf
                '
                Debug.Print "MAIL FROM:" & Trim$(txtSender)
                '
            Case MAIL_FROM
                '
                'Change current state of the session
                m_State = MAIL_RCPTTO
                '
                'Send RCPT TO command to the server
                Winsock1.SendData "RCPT TO:" & Trim$("dreamingweb@hotmail.com") & vbCrLf
                '
                Debug.Print "RCPT TO:" & Trim$("dreamingweb@hotmail.com")
                '
            Case MAIL_RCPTTO
                '
                'Change current state of the session
                m_State = MAIL_DATA
                '
                'Send DATA command to the server
                Winsock1.SendData "DATA" & vbCrLf
                '
                Debug.Print "DATA"
                '
            Case MAIL_DATA
                '
                'Change current state of the session
                m_State = MAIL_DOT
                '
                'So now we are sending a message body
                'Each line of text must be completed with
                'vbCrLf
                '
            
                Winsock1.SendData "From:" & txtSenderName & " <" & txtSender & ">" & vbCrLf
                Winsock1.SendData "To:" & "Danio Denirio" & " <" & "dreamingweb@hotmail.com" & ">" & vbCrLf
                
                Debug.Print "Subject: " & txtSubject
                '
                If Len(txtSender.Text) > 0 Then
                    Winsock1.SendData "Subject:" & txtSubject & vbCrLf
                    Winsock1.SendData "Reply-To:" & txtSenderName & " <" & txtSender & ">" & vbCrLf
                Else
                    Winsock1.SendData "Subject:" & txtSubject & vbCrLf
                End If
             
                Winsock1.SendData "Mime-Version: 1.0" & vbCrLf
                Winsock1.SendData "Content-Type: multipart/mixed; boundary=" & Chr$(34) & "NextMimePart" & Chr$(34) & vbCrLf
                Winsock1.SendData "Content-Transfer-Encoding: 7bit" & vbCrLf
                Winsock1.SendData "This is a multi-part message in MIME format." & vbCrLf
                Winsock1.SendData "--NextMimePart" & vbCrLf & "--NextMimePart" & vbCrLf & vbCrLf
                
                
                Dim strMessage As Variant
                Dim tDot As Long
                
                strMessage = txtMessage.Text & vbCrLf & vbCrLf & picgen.textcode.Text 'if its gone wronge its here Dreamingweb
                
                'The following routine is necessary in order to be able to send
                'a dot on a single line without confusing the serv:er
                '(Very Important, otherwhise the email might get truncated)

                tDot = 1
                For i = 1 To Len(strMessage)
                    tDot = InStr(tDot + 4, strMessage, vbCrLf & "." & vbCrLf)
                    If tDot = 0 Then Exit For
                    strMessage = Mid$(strMessage, 1, tDot + 2) & Chr$(0) & Mid$(strMessage, tDot + 3)
                    DoEvents
                Next
                

                Winsock1.SendData strMessage & vbCrLf
                strMessage = ""
                '
                'Send a dot symbol to inform server
                'that sending of message comleted
                Winsock1.SendData "." & vbCrLf
                '
                Debug.Print "."
                '
            Case MAIL_DOT
                'Change current state of the session
                m_State = MAIL_QUIT
                '
                'Send QUIT command to the server
                Winsock1.SendData "QUIT" & vbCrLf
                '
                Debug.Print "QUIT"
            Case MAIL_QUIT
                '
                'Close connection
                Winsock1.Close
                '
        End Select
       
    Else
        '
        'If we are here server replied with
        'unacceptable respose code therefore we need
        'close connection and inform user about problem
        '
        Winsock1.Close
        '

        If Not m_State = MAIL_QUIT Then
           Status.Caption = ("Status: " + "Error: " + strServerResponse)
        Else
            Status.Caption = ("Status: " + "Email Sent- Thanks For Your Emotion Picture")
        End If
        '
    End If
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  Status.Caption = "Error" & Number
End Sub

Private Sub Winsock1_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    Status.Caption = "Sending E-Mail..." + vbCrLf + vbCrLf + "(Bytes Remaining: " + Str(bytesRemaining) + ")"
End Sub
Private Sub SplitMessage(strMessage As String, strlines() As String)
Dim intAccs As Long
Dim i
Dim lngSpacePos As Long, lngStart As Long

    strMessage = Trim$(strMessage)
    lngSpacePos = 1
    lngSpacePos = InStr(lngSpacePos, strMessage, vbNewLine)
    
    Do While lngSpacePos
        intAccs = intAccs + 1
        lngSpacePos = InStr(lngSpacePos + 1, strMessage, vbNewLine)
    Loop
    
    ReDim strlines(intAccs)
    lngStart = 1
    
    For i = 0 To intAccs
        lngSpacePos = InStr(lngStart, strMessage, vbNewLine)
        
        If lngSpacePos Then
            strlines(i) = Mid(strMessage, lngStart, lngSpacePos - lngStart)
            lngStart = lngSpacePos + Len(vbNewLine)
        Else
            strlines(i) = Right(strMessage, Len(strMessage) - lngStart + 1)
        End If
    
    Next
End Sub
