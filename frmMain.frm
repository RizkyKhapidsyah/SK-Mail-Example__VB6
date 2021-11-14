VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Test POP3 connection"
   ClientHeight    =   4320
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   6300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQUIT 
      Caption         =   "QUIT"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3720
      TabIndex        =   10
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdDELE2 
      Caption         =   "DELE 2"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdDELE1 
      Caption         =   "DELE 1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1320
      TabIndex        =   8
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdRETR2 
      Caption         =   "RETR 2"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdRETR1 
      Caption         =   "RETR 1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4920
      TabIndex        =   6
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdLIST 
      Caption         =   "LIST"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdSTAT 
      Caption         =   "STAT"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdPASS 
      Caption         =   "PASS"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdUSER 
      Caption         =   "USER"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox txtCommunication 
      Height          =   3015
      Left            =   120
      Locked          =   -1  'True
      MaxLength       =   32000
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   6015
   End
   Begin MSWinsockLib.Winsock MailSock 
      Left            =   5835
      Top             =   90
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.Label Label2 
      Caption         =   "This sample was made by Frank H"
      Height          =   495
      Left            =   4920
      TabIndex        =   11
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   6405
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label1 
      Caption         =   "Communication:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.Menu Mnu_File 
      Caption         =   "&File"
      Begin VB.Menu Mnu_Connect 
         Caption         =   "&Connect&&Check"
         Enabled         =   0   'False
      End
      Begin VB.Menu Mnu_Setup 
         Caption         =   "&Setup"
      End
      Begin VB.Menu Mnu_Bindestreg 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_exit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdSend_Click()
  MailSock.SendData txtOutput.Text + NL
  AddText "you -> " + txtOutput.Text
  txtOutput.Text = ""
  txtOutput.SetFocus
End Sub


Private Sub cmdDELE1_Click()
MailSock.SendData "DELE 1" + NL
AddText "you -> DELE 1"
End Sub

Private Sub cmdDELE2_Click()
MailSock.SendData "DELE 2" + NL
AddText "you -> DELE 2"
End Sub

Private Sub cmdLIST_Click()
MailSock.SendData "LIST" + NL
AddText "you -> LIST"
End Sub

Private Sub cmdPASS_Click()
MailSock.SendData "PASS " + Password + NL
AddText "you -> PASS (password)"
End Sub

Private Sub cmdQUIT_Click()
MailSock.SendData "QUIT" + NL
AddText "you -> QUIT"
AddText "** Closing connection **"
End Sub

Private Sub cmdRETR1_Click()
MailSock.SendData "RETR 1" + NL
AddText "you -> RETR 1"
End Sub

Private Sub cmdRETR2_Click()
MailSock.SendData "RETR 2" + NL
AddText "you -> RETR 2"
End Sub

Private Sub cmdSTAT_Click()
MailSock.SendData "STAT" + NL
AddText "you -> STAT"
End Sub

Private Sub cmdUSER_Click()
MailSock.SendData "USER " + UserName + NL
AddText "you -> USER " + UserName
End Sub

Private Sub Form_Load()
  NL = Chr$(10) + Chr$(13)
End Sub

Private Sub MailSock_Close()
  AddText "** Connection Closed **"
End Sub

Private Sub MailSock_Connect()
  AddText "** Connection accepted **"
  cmdUSER.Enabled = True
  cmdPASS.Enabled = True
  cmdSTAT.Enabled = True
  cmdLIST.Enabled = True
  cmdRETR1.Enabled = True
  cmdRETR2.Enabled = True
  cmdDELE1.Enabled = True
  cmdDELE2.Enabled = True
  cmdQUIT.Enabled = True
End Sub

Private Sub MailSock_DataArrival(ByVal bytesTotal As Long)
  Dim InData As String
  MailSock.GetData InData
  AddText "server -> " + InData
End Sub

Private Sub Mnu_Connect_Click()
  MailSock.Protocol = sckTCPProtocol
  MailSock.RemotePort = 110          'default connect port
  MailSock.RemoteHost = ConnectTo    'set in SetUpForm
  MailSock.Close
  MailSock.Connect                   'Establish connection
  AddText "** Connecting to " + ConnectTo + " **"
End Sub

Private Sub Mnu_exit_Click()
  If MailSock.State <> sckClosed Then
    MailSock.Close
  End If
  End
End Sub

Private Sub Mnu_Setup_Click()
  frmSetUp.Show 1, Me
End Sub

Private Sub txtCommunication_Change()
  If Len(txtCommunication.Text) >= 31000 Then
    txtCommunication.Text = Mid$(txtCommunication.Text, Len(txtCommunication) - 31000, 31000)
  End If
End Sub

Public Sub AddText(Tekst As String)
  txtCommunication.Text = txtCommunication.Text + Tekst + Chr$(13) + Chr$(10)
  txtCommunication.SelStart = Len(txtCommunication.Text)
End Sub
