VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form MainForm 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   8250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   500
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   550
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Winsock 
      Index           =   0
      Left            =   120
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.PictureBox TailSprite 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   150
      Left            =   720
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox TailMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   150
      Left            =   960
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   720
      Top             =   3120
   End
   Begin VB.Timer MainTimer 
      Interval        =   50
      Left            =   120
      Top             =   120
   End
   Begin VB.Menu SetServer 
      Caption         =   "Server"
   End
   Begin VB.Menu SetClient 
      Caption         =   "Client"
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright © 2005 T-RonX Modding All rights reserved '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim StrBuffer() As String

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Updates ClientDirection
    DirectionKeys KeyCode
End Sub

Private Sub Form_Load()
'Temp.

JoinPort = 9609
JoinIP = "192.168.1.2"
ServerPort = 9609

    Me.Caption = "T-Snake 3  -  Build: " & App.Revision
    
    TailSprite.Picture = LoadPicture(App.Path & "\Tail_Sprite.bmp")
    TailMask.Picture = LoadPicture(App.Path & "\Tail_Mask.bmp")
    
    ClientTailCount = 3
    ReDim Preserve ClientInfo(0 To ClientTailCount)
    ClientInfo(0).Left = 250
    ClientInfo(1).Left = -10
    ClientInfo(2).Left = -10
    ClientInfo(3).Left = -10
    ClientDirection = "Down"

    OppTailCount = 3
    ReDim Preserve OppInfo(0 To OppTailCount)
    OppInfo(0).Left = 250
    OppInfo(1).Left = -10
    OppInfo(2).Left = -10
    OppInfo(3).Left = -10

End Sub

Private Sub MainTimer_Timer()

If Winsock(0).State = 7 Then
    Winsock(0).SendData CStr(ClientInfo(0).Top) & "." & CStr(ClientInfo(0).Left) & "." & CStr(ClientTailCount) & ","
End If

'Simply changes the position of the first tailpart in memory. The other parts and drawing are handled by RenderNextFrame.
Select Case ClientDirection
    Case "Right"
        ClientInfo(0).Left = ClientInfo(0).Left + 10
                                                 '^^
                                                 '10 is best for optimization, => 11 couses gfx errors.
                                                 'Tailpart cant move further ayway then its own size, else we get ugly gaps.
    Case "Left"
        ClientInfo(0).Left = ClientInfo(0).Left - 10
           
    Case "Up"
        ClientInfo(0).Top = ClientInfo(0).Top - 10

    Case "Down"
        ClientInfo(0).Top = ClientInfo(0).Top + 10
End Select
      
    Boarder_CollisionDetection 'Checks is snake crossed the boarder.
    Snake_CollisionDetection 'Checks if snake hits himself.
    RenderNextFrame
    

End Sub

Private Sub SetClient_Click()
    ConnectForm.Show vbModal
End Sub

Private Sub SetServer_Click()
    ServerForm.Show vbModal
End Sub

Private Sub Timer1_Timer()
    'Increases the size of the snake by 1 part. Just increases a number, rest is handled by RenderNextFrame
    
If Not ClientTailCount > 15 Then
    ClientTailCount = ClientTailCount + 1
    ReDim Preserve ClientInfo(0 To ClientTailCount)
End If

End Sub

Private Sub Winsock_Close(Index As Integer)
    If SckListening = True Then
        Server
        ReDim Preserve StrBuffer(0 To Winsock.UBound)   'Create buffer for user
    End If
End Sub

Private Sub Winsock_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    Winsock(Index).Close
    Winsock(Index).Accept requestID
    ReDim Preserve StrBuffer(0 To Winsock.UBound)   'Create buffer for user
    StrBuffer(Index) = ""                           'Clear buffer
End Sub

Private Sub Winsock_DataArrival(Index As Integer, ByVal bytesTotal As Long)

Dim strSplit() As String
Dim strComSplit() As String
Dim strData As String
Dim I As Long
strData = vbNullString
    
    Winsock(Index).GetData strData
    StrBuffer(Index) = StrBuffer(Index) & strData
    Me.Caption = StrBuffer(Index)
    If FindText(StrBuffer(Index), ",") = True Then
        strComSplit() = Split(StrBuffer(Index), ",")
        For I = LBound(strComSplit) To UBound(strComSplit) - 1
            If strComSplit(I) <> "" Then
                strSplit() = Split(strComSplit(I), ".")
                    
                OppInfo(0).Top = CInt(strSplit(0))
                OppInfo(0).Left = CInt(strSplit(1))
                OppTailCount = CInt(strSplit(2))
                    
                ReDim Preserve OppInfo(0 To OppTailCount)
            End If
        Next
    End If
End Sub
Function FindText(StrString As String, StringToFind As String) As Boolean
    If InStr(1, StrString, StringToFind, vbTextCompare) > 0 Then
        FindText = True
    Else
        FindText = False
    End If
End Function

