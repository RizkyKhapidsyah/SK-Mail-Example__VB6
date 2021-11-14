VERSION 5.00
Begin VB.Form frmSetUp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set up "
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3690
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "User information "
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   3375
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1200
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtUserName 
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   645
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   285
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Connection "
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.TextBox txtConnectTo 
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Text            =   "xxx.xxx.xxx.xxx or by name"
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Connect to:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   270
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmSetUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOK_Click()
  ConnectTo = txtConnectTo.Text
  UserName = txtUserName.Text
  Password = txtPassword.Text
  frmMain.Mnu_Connect.Enabled = True
  Unload Me
End Sub

Private Sub txtConnectTo_gotfocus()
  txtConnectTo.SelStart = 0
  txtConnectTo.SelLength = Len(txtConnectTo.Text)
End Sub

Private Sub txtPassword_gotfocus()
  txtPassword.SelStart = 0
  txtPassword.SelLength = Len(txtPassword.Text)
End Sub

Private Sub txtUserName_gotfocus()
  txtUserName.SelStart = 0
  txtUserName.SelLength = Len(txtUserName.Text)
End Sub
