Attribute VB_Name = "modEngine"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright © 2005 T-RonX Modding All rights reserved '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

 'Declares teh BitBlt variable.
 Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
 Public Const SRCPAINT = &HEE0086
 Public Const SRCAND = &H8800C6
 Public SckListening As Boolean
 
Public Function RenderNextFrame()
'We are rendering the same pic to an other place using the For statement, so you actually see the same pic but it looks like many pics, we are actually fooling Windows. I don't fully understand it myself :/  but it works... kinda, whatever. To help you understand, try to temporarily remove MainForm.Cls.
'Cleans up the form.
MainForm.Cls
    
   'We go from top to bottom so Tail(i) can get the coördinates of the tailpart thats in from of him.
   For I = ClientTailCount To 1 Step -1

        '0 - 1(i - 1) is not a possable.
        If Not I = 0 Then
            BitBlt MainForm.hDC, ClientInfo(I - 1).Left, ClientInfo(I - 1).Top, 10, 10, MainForm.TailMask.hDC, 0, 0, SRCAND 'First we need to render the mask.
            BitBlt MainForm.hDC, ClientInfo(I - 1).Left, ClientInfo(I - 1).Top, 10, 10, MainForm.TailSprite.hDC, 0, 0, SRCPAINT 'Then the sprite.
        End If
        
        'Updates coördinates.
        ClientInfo(I).Left = ClientInfo(I - 1).Left
        ClientInfo(I).Top = ClientInfo(I - 1).Top
   Next

   For I = OppTailCount To 1 Step -1
        
        '0 - 1(i - 1) is not a possable.
        If Not I = 0 Then
            BitBlt MainForm.hDC, OppInfo(I - 1).Left, OppInfo(I - 1).Top, 10, 10, MainForm.TailMask.hDC, 0, 0, SRCAND 'First we need to render the mask.
            BitBlt MainForm.hDC, OppInfo(I - 1).Left, OppInfo(I - 1).Top, 10, 10, MainForm.TailSprite.hDC, 0, 0, SRCPAINT 'Then the sprite.
        End If
        
        'Updates coördinates.
        OppInfo(I).Left = OppInfo(I - 1).Left
        OppInfo(I).Top = OppInfo(I - 1).Top

   Next

'Redraws the form.
MainForm.Refresh

End Function

Public Function Snake_CollisionDetection()

'Checks if the first tail part has hit any other
For I = 1 To ClientTailCount
        If ClientInfo(0).Top + 9 >= ClientInfo(I).Top And ClientInfo(0).Top <= ClientInfo(I).Top + 9 And ClientInfo(0).Left + 9 >= ClientInfo(I).Left And ClientInfo(0).Left <= ClientInfo(I).Left + 9 Then
        End If
Next

End Function

Public Function Boarder_CollisionDetection()

        'Snake crossed top.
        If ClientInfo(0).Top < 0 Then
            ClientInfo(0).Top = 490 'Makes it appear on the other side of the thingy.
            ClientDirection = "Up"
        End If
        
        'Snake crossed bottom.
        If ClientInfo(0).Top > MainForm.ScaleHeight - 10 Then
            ClientInfo(0).Top = 0
            ClientDirection = "Down"
        End If
        
        'Snake crossed left.
        If ClientInfo(0).Left < 0 Then
            ClientInfo(0).Left = 540
            ClientDirection = "Left"
        End If
        
        'Snake crossed right.
        If ClientInfo(0).Left > MainForm.ScaleWidth - 10 Then
            ClientInfo(0).Left = 0
            ClientDirection = "Right"
        End If

End Function

Public Function Server()
    MainForm.Winsock(0).Close
    MainForm.Winsock(0).LocalPort = ServerPort
    MainForm.Winsock(0).Listen
    SckListening = True
End Function
