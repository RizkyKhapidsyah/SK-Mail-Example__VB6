VERSION 5.00
Begin VB.Form frmSetUp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set Up"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   3360
      TabIndex        =   12
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Message"
      Height          =   2535
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   4455
      Begin VB.TextBox txtMessage 
         Height          =   1575
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   840
         Width           =   4215
      End
      Begin VB.TextBox txtRecepient 
         Height          =   285
         Left            =   1200
         TabIndex        =   10
         Text            =   "email@your.recepient.isp"
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label5 
         Caption         =   "Message:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Recepient:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   300
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Connection"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txtSender 
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Text            =   "email@your.isp your name (ie: test@test.com (My name))"
         Top             =   960
         Width           =   3135
      End
      Begin VB.TextBox txtSystemName 
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Text            =   "Something (ie: JohnDoe)"
         Top             =   600
         Width           =   3135
      End
      Begin VB.TextBox txtSMTP 
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Text            =   "your.mail.isp"
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label3 
         Caption         =   "Sender e-mail:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1005
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Systemname:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   645
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "SMTP-server:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   300
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmSetUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
  Message = txtMessage.Text
  Recepient = txtRecepient.Text
  Sender = txtSender.Text
  SystemName = txtSystemName.Text
  SMTPServer = txtSMTP.Text
  FrmMain.Mnu_Connect.Enabled = True
  Unload Me
End Sub

Private Sub txtMessage_gotfocus()
  txtMessage.SelStart = 0
  txtMessage.SelLength = Len(txtMessage.Text)
End Sub

Private Sub txtRecepient_gotfocus()
  txtRecepient.SelStart = 0
  txtRecepient.SelLength = Len(txtRecepient.Text)
End Sub

Private Sub txtSender_gotfocus()
  txtSender.SelStart = 0
  txtSender.SelLength = Len(txtSender.Text)
End Sub

Private Sub txtSMTP_GotFocus()
  txtSMTP.SelStart = 0
  txtSMTP.SelLength = Len(txtSMTP.Text)
End Sub

Private Sub txtSystemName_gotfocus()
  txtSystemName.SelStart = 0
  txtSystemName.SelLength = Len(txtSystemName.Text)
End Sub
