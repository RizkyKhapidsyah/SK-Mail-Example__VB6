VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Test SMTP connection"
   ClientHeight    =   4215
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQUIT 
      Caption         =   "QUIT"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   7
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdSENDMSG 
      Caption         =   "SEND MSG"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      ToolTipText     =   "When finished with the msg, send following keystroke: chr$(13)+chr$(10)+"".""+chr$(13)+chr$(10) (This is done auto. in this prg.)"
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdRCPTTO 
      Caption         =   "RCPT TO"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdMAILFROM 
      Caption         =   "MAIL FROM"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdRSET 
      Caption         =   "RSET"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdEHLO 
      Caption         =   "EHLO"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtCommunication 
      Height          =   2895
      Left            =   120
      MaxLength       =   32000
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   360
      Width           =   5775
   End
   Begin MSWinsockLib.Winsock MailSock 
      Left            =   5280
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.Line Line1 
      X1              =   6000
      X2              =   0
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label1 
      Caption         =   "Communication:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Menu Mnu_Filer 
      Caption         =   "&File"
      Begin VB.Menu Mnu_Connect 
         Caption         =   "&Connect&&Send"
         Enabled         =   0   'False
      End
      Begin VB.Menu Mnu_SetUp 
         Caption         =   "&Set up"
      End
      Begin VB.Menu Mnu_Bindestreg 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_Exit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEHLO_Click()
  MailSock.SendData "HELO " + SystemName + NL
  AddText "you -> HELO " + SystemName
End Sub

Private Sub cmdMAILFROM_Click()
  MailSock.SendData "MAIL FROM:<" + Sender + ">" + NL
  AddText "you -> MAIL FROM:<" + Sender + ">"
End Sub

Private Sub cmdQUIT_Click()
  MailSock.SendData "QUIT" + NL
  AddText "you -> QUIT"
  AddText "** Closing connection **"
End Sub

Private Sub cmdRCPTTO_Click()
  MailSock.SendData "RCPT TO:" + Recepient + NL
  AddText "RCPT TO:" + Recepient
End Sub

Private Sub cmdRSET_Click()
  MailSock.SendData "RSET" + NL
  AddText "you -> RSET"
End Sub

Private Sub cmdSENDMSG_Click()
  MailSock.SendData "DATA" + NL
  AddText "you -> DATA"
  MailSock.SendData Message + NL + "." + NL
  AddText "you -> " + Message + NL + "."
End Sub

Private Sub Form_Load()
    
    NL = Chr$(13) + Chr$(10)
End Sub

Private Sub MailSock_Close()
  AddText "** Connection Closed **"
End Sub

Private Sub MailSock_Connect()
  AddText "** Connection accepted **"
  cmdQUIT.Enabled = True
  cmdEHLO.Enabled = True
  cmdRSET.Enabled = True
  cmdMAILFROM.Enabled = True
  cmdRCPTTO.Enabled = True
  cmdSENDMSG.Enabled = True
End Sub

Private Sub MailSock_DataArrival(ByVal bytesTotal As Long)
  Dim InData As String
  MailSock.GetData InData
  AddText "server -> " + InData
End Sub

Private Sub Mnu_Connect_Click()
  If MailSock.State <> sckClosed Then
    MailSock.Close
  End If
  MailSock.Protocol = sckTCPProtocol
  MailSock.RemotePort = 25
  MailSock.RemoteHost = SMTPServer
  MailSock.Connect
  AddText "** Connecting to " + SMTPServer + " **"
End Sub

Private Sub Mnu_Exit_Click()
  If MailSock.State <> sckClosed Then
    MailSock.Close
  End If
  End
End Sub

Private Sub Mnu_SetUp_Click()
  frmSetUp.Show vbModal
End Sub

Public Sub AddText(Tekst As String)
  txtCommunication.Text = txtCommunication.Text + Tekst + Chr$(13) + Chr$(10)
  txtCommunication.SelStart = Len(txtCommunication.Text)
End Sub

Private Sub txtCommunication_Change()
  If Len(txtCommunication.Text) >= 31000 Then
    txtCommunication.Text = Mid$(txtCommunication.Text, Len(txtCommunication) - 31000, 31000)
  End If
End Sub
