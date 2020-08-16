VERSION 5.00
Begin VB.Form ConnectForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Connect"
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   2040
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton ConnectForm 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox TextPort 
      Height          =   285
      Left            =   480
      TabIndex        =   3
      Text            =   "9609"
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox TextIP 
      Height          =   285
      Left            =   480
      TabIndex        =   0
      Text            =   "192.168.1.2"
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Port:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "IP:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
End
Attribute VB_Name = "ConnectForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ConnectForm_Click()
    SckListening = False
    JoinPort = CInt(TextPort.Text)
    JoinIP = TextIP.Text
    
    MainForm.Winsock(0).Close
    MainForm.Winsock(0).RemoteHost = JoinIP
    MainForm.Winsock(0).RemotePort = JoinPort
    MainForm.Winsock(0).Connect
    
    Unload Me
End Sub
