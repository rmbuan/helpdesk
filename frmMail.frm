VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMail 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SMTP E-Mail"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7140
   Icon            =   "frmMail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Send E-Mail"
      Height          =   495
      Left            =   960
      TabIndex        =   10
      Top             =   4440
      Width           =   2175
   End
   Begin VB.TextBox txtFromEmailAddress 
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   3255
   End
   Begin VB.TextBox txtToEmailAddress 
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   3255
   End
   Begin VB.TextBox txtEmailSubject 
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   3255
   End
   Begin VB.TextBox txtEmailBodyOfMessage 
      Height          =   1455
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   2160
      Width           =   6855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   4440
      TabIndex        =   5
      Top             =   4440
      Width           =   1695
   End
   Begin VB.TextBox txtFromName 
      Height          =   285
      Left            =   3600
      TabIndex        =   4
      Top             =   480
      Width           =   3375
   End
   Begin VB.TextBox ToNametxt 
      Height          =   285
      Left            =   3600
      TabIndex        =   3
      Top             =   1080
      Width           =   3375
   End
   Begin VB.TextBox txtEmailServer 
      Height          =   285
      Left            =   3600
      TabIndex        =   2
      Text            =   "smail.gbrownc.on.ca"
      Top             =   1680
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      Caption         =   "Status:"
      Height          =   615
      Left            =   960
      TabIndex        =   0
      Top             =   3720
      Width           =   5175
      Begin VB.Label StatusTxt 
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4935
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3480
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "From (e-mail address)"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "To"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Subject"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Your Name"
      Height          =   255
      Left            =   3600
      TabIndex        =   13
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label Label5 
      Caption         =   "Recipient Name"
      Height          =   255
      Left            =   3600
      TabIndex        =   12
      Top             =   840
      Width           =   3375
   End
   Begin VB.Label Label6 
      Caption         =   "E-Mail Server"
      Height          =   255
      Left            =   3600
      TabIndex        =   11
      Top             =   1440
      Width           =   3375
   End
End
Attribute VB_Name = "frmMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit ' Set variables explicitly

Dim Response As String, Reply As Integer, DateNow As String
Dim first As String, Second As String, Third As String
Dim Fourth As String, Fifth As String, Sixth As String
Dim Seventh As String, Eighth As String, Ninth As String
Dim Start As Single, Tmr As Single, MsgTitle As String



Sub SendEmail(MailServerName As String, FromName As String, FromEmailAddress As String, ToName As String, ToEmailAddress As String, EmailSubject As String, EmailBodyOfMessage As String)
          
    Winsock1.LocalPort = 0 ' Must set local port to 0 (Zero) or you can only send 1 e-mail pre program start
    
If Winsock1.State = sckClosed Then ' Check to see if socet is closed
    DateNow = Format(Date, "Ddd") & ", " & Format(Date, "dd Mmm YYYY") & " " & Format(Time, "hh:mm:ss") & "" & " -0600"
    first = "mail from:" + Chr(32) + FromEmailAddress + vbCrLf ' Get who's sending E-Mail address
    Second = "rcpt to:" + Chr(32) + ToEmailAddress + vbCrLf ' Get who mail is going to
    Third = "Date:" + Chr(32) + DateNow + vbCrLf ' Date when being sent
    Fourth = "From:" + Chr(32) + FromName + vbCrLf ' Who's Sending
    Fifth = "To:" + Chr(32) + ToNametxt + vbCrLf ' Who it going to
    Sixth = "Subject:" + Chr(32) + EmailSubject + vbCrLf ' Subject of E-Mail
    Seventh = EmailBodyOfMessage + vbCrLf ' E-mail message body
    Ninth = "X-Mailer: EBT Reporter v 2.x" + vbCrLf ' What program sent the e-mail, customize this
    Eighth = Fourth + Third + Ninth + Fifth + Sixth  ' Combine for proper SMTP sending

    Winsock1.Protocol = sckTCPProtocol ' Set protocol for sending
    Winsock1.RemoteHost = MailServerName ' Set the server address
    Winsock1.RemotePort = 25 ' Set the SMTP Port
    Winsock1.Connect ' Start connection
    
    WaitFor ("220")
    
    StatusTxt.Caption = "Connecting...."
    StatusTxt.Refresh
    
    Winsock1.SendData ("HELO worldcomputers.com" + vbCrLf)

    WaitFor ("250")

    StatusTxt.Caption = "Connected"
    StatusTxt.Refresh

    Winsock1.SendData (first)

    StatusTxt.Caption = "Sending Message"
    StatusTxt.Refresh

    WaitFor ("250")

    Winsock1.SendData (Second)

    WaitFor ("250")

    Winsock1.SendData ("data" + vbCrLf)
    
    WaitFor ("354")


    Winsock1.SendData (Eighth + vbCrLf)
    Winsock1.SendData (Seventh + vbCrLf)
    Winsock1.SendData ("." + vbCrLf)

    WaitFor ("250")

    Winsock1.SendData ("quit" + vbCrLf)
    
    StatusTxt.Caption = "Disconnecting"
    StatusTxt.Refresh

    WaitFor ("221")

    Winsock1.Close
Else
    MsgBox (Str(Winsock1.State))
End If
   
End Sub

Sub WaitFor(ResponseCode As String)
    Start = Timer ' Time event so won't get stuck in loop
    While Len(Response) = 0
        Tmr = Start - Timer
        DoEvents ' Let System keep checking for incoming response **IMPORTANT**
        If Tmr > 50 Then ' Time in seconds to wait
            MsgBox "SMTP service error, timed out while waiting for response", 64, MsgTitle
            Exit Sub
        End If
    Wend
    While Left(Response, 3) <> ResponseCode
        DoEvents
        If Tmr > 50 Then
            MsgBox "SMTP service error, impromper response code. Code should have been: " + ResponseCode + " Code recieved: " + Response, 64, MsgTitle
            Exit Sub
        End If
    Wend
Response = "" ' Sent response code to blank **IMPORTANT**
End Sub

Private Sub Command1_Click()
    SendEmail txtEmailServer.Text, txtFromName.Text, txtFromEmailAddress.Text, txtToEmailAddress.Text, txtToEmailAddress.Text, txtEmailSubject.Text, txtEmailBodyOfMessage.Text
    'MsgBox ("Mail Sent")
    StatusTxt.Caption = "Mail Sent"
    StatusTxt.Refresh
    Beep
    
    Close
End Sub

Private Sub Command2_Click()
    
    Unload Me
    
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

    Winsock1.GetData Response ' Check for incoming response *IMPORTANT*

End Sub

