Attribute VB_Name = "modAdditionalCode"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright © 2005 T-RonX Modding All rights reserved '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function DirectionKeys(KeyCode)
'Changes Direction variable so we know where the snake is moving to. The actual movement is handles by the main timer.
Select Case KeyCode
        Case vbKeyLeft
            If ClientDirection = "Right" Then
            Else
                ClientDirection = "Left"
            End If
            
        Case vbKeyRight
            If ClientDirection = "Left" Then
            Else
                ClientDirection = "Right"
            End If
            
        Case vbKeyUp
            If ClientDirection = "Down" Then
            Else
                ClientDirection = "Up"
            End If

        Case vbKeyDown
            If ClientDirection = "Up" Then
            Else
                ClientDirection = "Down"
            End If

        Case vbKeyA
            If ClientDirection = "Right" Then
            Else
                ClientDirection = "Left"
            End If
            
        Case vbKeyD
            If ClientDirection = "Left" Then
            Else
                ClientDirection = "Right"
            End If
            
        Case vbKeyW
            If ClientDirection = "Down" Then
            Else
                ClientDirection = "Up"
            End If

        Case vbKeyS
            If ClientDirection = "Up" Then
            Else
                ClientDirection = "Down"
            End If

End Select

End Function

Public Function GetState(pSock As Winsock) As String
   'Return a string representation of the state.
   Select Case pSock.State
     Case 0
       GetState = "Closed"
     Case 1
       GetState = "Open"
     Case 2
       GetState = "Listening"
     Case 3
       GetState = "Connection Pending"
     Case 4
       GetState = "Resolving Host"
     Case 5
       GetState = "Host Resolved"
     Case 6
       GetState = "Connecting"
     Case 7
       GetState = "Connected"
     Case 8
       GetState = "Closing"
     Case 9
       GetState = "Error"
   End Select
End Function

