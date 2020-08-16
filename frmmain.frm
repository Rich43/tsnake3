VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   500
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   500
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer MainSequenceTimer 
      Interval        =   80
      Left            =   6960
      Top             =   120
   End
   Begin VB.PictureBox ClientSP 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H0015559B&
      ForeColor       =   &H80000008&
      Height          =   150
      Index           =   0
      Left            =   3120
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   0
      Top             =   1440
      Width           =   150
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ClientSP_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    InputHandler KeyCode
End Sub

Private Sub Form_Load()
    Initialize
End Sub

Private Sub MainSequenceTimer_Timer()
    MainSequence
End Sub

Public Sub Initialize()
    Me.Caption = "TS Engine 3 - Build: " & App.Revision
    c_Direction = Dir.Down
    
    For i = 1 To 9
        Load ClientSP(i)
        ClientSP(i).Visible = True
    Next

End Sub

