Attribute VB_Name = "modGameEngine"

Public Sub MainSequence()

    Select Case c_Direction
        Case Dir.Up
            frmMain.ClientSP(0).Top = frmMain.ClientSP(0).Top - 10
            For i = frmMain.ClientSP.UBound To 1 Step -1
                frmMain.ClientSP(i).Top = frmMain.ClientSP(i - 1).Top
                frmMain.ClientSP(i).Left = frmMain.ClientSP(i - 1).Left
            Next
        
        Case Dir.Down
            frmMain.ClientSP(0).Top = frmMain.ClientSP(0).Top + 10
            For i = frmMain.ClientSP.UBound To 1 Step -1
                frmMain.ClientSP(i).Top = frmMain.ClientSP(i - 1).Top
                frmMain.ClientSP(i).Left = frmMain.ClientSP(i - 1).Left
            Next
        
        Case Dir.Left
            frmMain.ClientSP(0).Left = frmMain.ClientSP(0).Left - 10
            For i = frmMain.ClientSP.UBound To 1 Step -1
                frmMain.ClientSP(i).Top = frmMain.ClientSP(i - 1).Top
                frmMain.ClientSP(i).Left = frmMain.ClientSP(i - 1).Left
            Next
        
        Case Dir.Right
            frmMain.ClientSP(0).Left = frmMain.ClientSP(0).Left + 10
            For i = frmMain.ClientSP.UBound To 1 Step -1
                frmMain.ClientSP(i).Top = frmMain.ClientSP(i - 1).Top
                frmMain.ClientSP(i).Left = frmMain.ClientSP(i - 1).Left
            Next
        
    End Select

End Sub

Public Sub InputHandler(KeyCode As Integer)

    Select Case KeyCode
        Case vbKeyUp
            c_Direction = Dir.Up
            
        Case vbKeyDown
            c_Direction = Dir.Down
            
        Case vbKeyLeft
            c_Direction = Dir.Left
            
        Case vbKeyRight
            c_Direction = Dir.Right
            
    End Select
    
End Sub
