VERSION 5.00
Begin VB.Form ServerForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   2280
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox TextPort 
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Text            =   "9609"
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Port:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   375
   End
End
Attribute VB_Name = "ServerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    ServerPort = CInt(TextPort.Text)
    Server
    Unload Me
End Sub
