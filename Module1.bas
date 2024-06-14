Attribute VB_Name = "Module1"
Option Explicit

Public Mode As Integer
Dim Protocol As Long
Dim Score As Integer
'Packman
Public R As Long, C As Integer
'Enemy
Public R1 As Long, C1 As Integer, Z1 As String * 1
Public R2 As Long, C2 As Integer, Z2 As String * 1
Public R3 As Long, C3 As Integer, Z3 As String * 1
Public R4 As Long, C4 As Integer, Z4 As String * 1

Public MR1 As Long, MC1 As Integer, WR1 As Long, WC1 As Integer
Public MR2 As Long, MC2 As Integer, WR2 As Long, WC2 As Integer
Public MR3 As Long, MC3 As Integer, WR3 As Long, WC3 As Integer
Public MR4 As Long, MC4 As Integer, WR4 As Long, WC4 As Integer
Public Sub Right()
    If Mode = 0 Then
            Cells(Protocol, 80) = "R"
            Protocol = Protocol + 1
        If Cells(R, C + 1).Interior.ColorIndex = 6 And Cells(R, C + 1) <> "L" Then
            Cells(R, C) = ""
            C = C + 1
            If C = 28 Then C = 2
            If Cells(R, C) = Chr(159) Then
                Score = Score + 1
                [AF2] = Score
            End If
            Cells(R, C) = "J": Cells(R, C).Font.ColorIndex = 3
        End If
        If Score = 300 Then
        MsgBox "Winner"
        End If
        Enemy
    End If
End Sub
Public Sub RightReplay()
    If Cells(R, C + 1).Interior.ColorIndex = 6 And Cells(R, C + 1) <> "L" Then
        Cells(R, C) = ""
        C = C + 1
        If C = 28 Then C = 2
        If Cells(R, C) = Chr(159) Then
            Score = Score + 1
            [AF2] = Score
        End If
        Cells(R, C) = "J": Cells(R, C).Font.ColorIndex = 3
    End If
    Enemy
End Sub
Public Sub Left()
    If Mode = 0 Then
            Cells(Protocol, 80) = "L"
            Protocol = Protocol + 1
        If Cells(R, C - 1).Interior.ColorIndex = 6 And Cells(R, C - 1) <> "L" Then
            Cells(R, C) = ""
            C = C - 1
            If C = 1 Then C = 27
            If Cells(R, C) = Chr(159) Then
                Score = Score + 1
                [AF2] = Score
            End If
            Cells(R, C) = "J": Cells(R, C).Font.ColorIndex = 3
        End If
        If Score = 300 Then
        MsgBox "Winner"
        End If
        Enemy
    End If
End Sub
Public Sub LeftReplay()
    If Cells(R, C - 1).Interior.ColorIndex = 6 And Cells(R, C - 1) <> "L" Then
        Cells(R, C) = ""
        C = C - 1
        If C = 1 Then C = 27
        If Cells(R, C) = Chr(159) Then
            Score = Score + 1
            [AF2] = Score
        End If
        Cells(R, C) = "J": Cells(R, C).Font.ColorIndex = 3
    End If
    Enemy
End Sub
Public Sub Up()
    If Mode = 0 Then
            Cells(Protocol, 80) = "U"
            Protocol = Protocol + 1
        If Cells(R - 1, C).Interior.ColorIndex = 6 And Cells(R - 1, C) <> "L" Then
            Cells(R, C) = ""
            R = R - 1
            If Cells(R, C) = Chr(159) Then
                Score = Score + 1
                [AF2] = Score
            End If
            Cells(R, C) = "J": Cells(R, C).Font.ColorIndex = 3
        End If
        If Score = 300 Then
        MsgBox "Winner"
        End If
        Enemy
    End If
End Sub
Public Sub UpReplay()
    If Cells(R - 1, C).Interior.ColorIndex = 6 And Cells(R - 1, C) <> "L" Then
        Cells(R, C) = ""
        R = R - 1
        If Cells(R, C) = Chr(159) Then
            Score = Score + 1
            [AF2] = Score
        End If
        Cells(R, C) = "J": Cells(R, C).Font.ColorIndex = 3
    End If
    Enemy
End Sub
Public Sub Down()
    If Mode = 0 Then
        Cells(Protocol, 80) = "D"
        Protocol = Protocol + 1
        If Cells(R + 1, C).Interior.ColorIndex = 6 And Cells(R + 1, C) <> "L" Then
            Cells(R, C) = ""
            R = R + 1
            If Cells(R, C) = Chr(159) Then
                Score = Score + 1
                [AF2] = Score
            End If
            Cells(R, C) = "J": Cells(R, C).Font.ColorIndex = 3
        End If
        If Score = 300 Then
        MsgBox "Winner"
        End If
        Enemy
    End If
End Sub
Public Sub DownReplay()
    If Cells(R + 1, C).Interior.ColorIndex = 6 And Cells(R + 1, C) <> "L" Then
        Cells(R, C) = ""
        R = R + 1
        If Cells(R, C) = Chr(159) Then
            Score = Score + 1
            [AF2] = Score
        End If
        Cells(R, C) = "J": Cells(R, C).Font.ColorIndex = 3
    End If
    Enemy
End Sub
Public Sub Enemy()
    Dim DR1 As Long, DC1 As Integer
    Dim DR2 As Long, DC2 As Integer
    Dim DR3 As Long, DC3 As Integer
    Dim DR4 As Long, DC4 As Integer
    If (MR1 <> 0 Or MC1 <> 0) Then
        If Cells(R1 + WR1, C1 + WC1).Interior.ColorIndex = 6 And Cells(R1 + WR1, C1 + WC1) <> "L" Then ' Turn
            DR1 = WR1: DC1 = WC1
            MR1 = 0: MC1 = 0
        ElseIf Cells(R1 + MR1, C1 + MC1).Interior.ColorIndex = 6 And Cells(R1 + WR1, C1 + MC1) <> "L" Then ' From dead end to turn
            DR1 = MR1: DC1 = MC1
        Else
            DR1 = 0: DC1 = 0
            MR1 = 0: MC1 = 0
        End If
    ElseIf Abs(R - R1) > Abs(C - C1) Then ' Prefered vertical movement
        If R < R1 And Cells(R1 - 1, C1).Interior.ColorIndex = 6 And Cells(R1 - 1, C1) <> "L" Then ' Up
            DR1 = -1: DC1 = 0
        ElseIf R > R1 And Cells(R1 + 1, C1).Interior.ColorIndex = 6 And Cells(R1 + 1, C1) <> "L" Then ' Down
            DR1 = 1: DC1 = 0
        ElseIf C < C1 And Cells(R1, C1 - 1).Interior.ColorIndex = 6 And Cells(R1, C1 - 1) <> "L" Then ' Left
            DR1 = 0: DC1 = -1
        ElseIf C > C1 And Cells(R1, C1 + 1).Interior.ColorIndex = 6 And Cells(R1, C1 + 1) <> "L" Then ' Right
            DR1 = 0: DC1 = 1
        ElseIf Cells(R1, C1 - 1).Interior.ColorIndex = 6 And Cells(R1, C1 - 1) <> "L" Then
            DR1 = 0: DC1 = -1
            MR1 = 0: MC1 = -1: WR1 = Sgn(R - R1): WC1 = 0
        ElseIf Cells(R1, C1 + 1).Interior.ColorIndex = 6 And Cells(R1, C1 + 1) <> "L" Then
            DR1 = 0: DC1 = 1
            MR1 = 0: MC1 = 1: WR1 = Sgn(R - R1): WC1 = 0
        End If

    Else ' Prefered horizontal movement
        If C < C1 And Cells(R1, C1 - 1).Interior.ColorIndex = 6 And Cells(R1, C1 - 1) <> "L" Then ' Left
            DR1 = 0: DC1 = -1
        ElseIf C > C1 And Cells(R1, C1 + 1).Interior.ColorIndex = 6 And Cells(R1, C1 + 1) <> "L" Then ' Right
            DR1 = 0: DC1 = 1
        ElseIf R < R1 And Cells(R1 - 1, C1).Interior.ColorIndex = 6 And Cells(R1 - 1, C1) <> "L" Then ' Up
            DR1 = -1: DC1 = 0
        ElseIf R > R1 And Cells(R1 + 1, C1).Interior.ColorIndex = 6 And Cells(R1 + 1, C1) <> "L" Then ' Down
            DR1 = 1: DC1 = 0
        ElseIf Cells(R1 - 1, C1).Interior.ColorIndex = 6 And Cells(R1 - 1, C1) <> "L" Then
            DR1 = -1: DC1 = 0
            MR1 = -1: MC1 = 0: WR1 = 0: WC1 = Sgn(C - C1)
        ElseIf Cells(R1 + 1, C1).Interior.ColorIndex = 6 And Cells(R1 + 1, C1) <> "L" Then
            DR1 = 1: DC1 = 0
            MR1 = 1: MC1 = 0: WR1 = 0: WC1 = Sgn(C - C1)
        End If
    End If
      If Cells(R1 + DR1, C1 + DC1) <> "L" Then
        Cells(R1, C1) = Z1
        Cells(R1, C1).Font.ColorIndex = 1
        C1 = C1 + DC1: R1 = R1 + DR1
        Z1 = Cells(R1, C1)
        Cells(R1, C1) = "L"
        Cells(R1, C1).Font.ColorIndex = 4
      End If
      If R1 = R And C1 = C Then
            MsgBox "Game Over"
            Begin
      End If
'Enemy 2
If (MR2 <> 0 Or MC2 <> 0) Then
        If Cells(R2 + WR2, C2 + WC2).Interior.ColorIndex = 6 And Cells(R2 + WR2, C2 + WC2) <> "L" Then ' Turn
            DR2 = WR2: DC2 = WC2
            MR2 = 0: MC2 = 0
        ElseIf Cells(R2 + MR2, C2 + MC2).Interior.ColorIndex = 6 And Cells(R2 + MR2, C2 + MC2) <> "L" Then ' From dead end to turn
            DR2 = MR2: DC2 = MC2
        Else
            DR2 = 0: DC2 = 0
            MR2 = 0: MC2 = 0
        End If
    ElseIf Abs(R - R2) > Abs(C - C2) Then ' Prefered vertical movement
        If R < R2 And Cells(R2 - 1, C2).Interior.ColorIndex = 6 And Cells(R2 - 1, C2) <> "L" Then ' Up
            DR2 = -1: DC2 = 0
        ElseIf R > R2 And Cells(R2 + 1, C2).Interior.ColorIndex = 6 And Cells(R2 + 1, C2) <> "L" Then ' Down
            DR2 = 1: DC2 = 0
        ElseIf C < C2 And Cells(R2, C2 - 1).Interior.ColorIndex = 6 And Cells(R2, C2 - 1) <> "L" Then ' Left
            DR2 = 0: DC2 = -1
        ElseIf C > C2 And Cells(R2, C2 + 1).Interior.ColorIndex = 6 And Cells(R2, C2 + 1) <> "L" Then ' Right
            DR2 = 0: DC2 = 1

        ElseIf Cells(R2, C2 - 1).Interior.ColorIndex = 6 And Cells(R2, C2 - 1) <> "L" Then
            DR2 = 0: DC2 = -1
            MR2 = 0: MC2 = -1: WR2 = Sgn(R - R2): WC2 = 0
        ElseIf Cells(R2, C2 + 1).Interior.ColorIndex = 6 And Cells(R2, C2 + 1) <> "L" Then
            DR2 = 0: DC2 = 1
            MR2 = 0: MC2 = 1: WR2 = Sgn(R - R2): WC2 = 0
        End If

    Else ' Prefered horizontal movement
        If C < C2 And Cells(R2, C2 - 1).Interior.ColorIndex = 6 And Cells(R2, C2 - 1) <> "L" Then ' Left
            DR2 = 0: DC2 = -1
        ElseIf C > C2 And Cells(R2, C2 + 1).Interior.ColorIndex = 6 And Cells(R2, C2 + 1) <> "L" Then ' Right
            DR2 = 0: DC2 = 1
        ElseIf R < R2 And Cells(R2 - 1, C2).Interior.ColorIndex = 6 And Cells(R2 - 1, C2) <> "L" Then ' Up
            DR2 = -1: DC2 = 0
        ElseIf R > R2 And Cells(R2 + 1, C2).Interior.ColorIndex = 6 And Cells(R2 + 1, C2) <> "L" Then ' Down
            DR2 = 1: DC2 = 0

        ElseIf Cells(R2 - 1, C2).Interior.ColorIndex = 6 And Cells(R2 - 1, C2) <> "L" Then
            DR2 = -1: DC2 = 0
            MR2 = -1: MC2 = 0: WR2 = 0: WC2 = Sgn(C - C2)
        ElseIf Cells(R2 + 1, C2).Interior.ColorIndex = 6 And Cells(R2 + 1, C2) <> "L" Then
            DR2 = 1: DC2 = 0
            MR2 = 1: MC2 = 0: WR2 = 0: WC2 = Sgn(C - C2)
        End If
    End If
    If Cells(R2 + DR2, C2 + DC2) <> "L" Then
        Cells(R2, C2) = Z2
        Cells(R2, C2).Font.ColorIndex = 1
        C2 = C2 + DC2: R2 = R2 + DR2
        Z2 = Cells(R2, C2)
        Cells(R2, C2) = "L"
        Cells(R2, C2).Font.ColorIndex = 4
     End If
     If R2 = R And C2 = C Then
        MsgBox "Game Over"
        Start
     End If
End Sub
    

' End Sub
' Row - horizontal
' Column - vertical

Public Sub Begin()
    Dim T As String, D As Single
    [A1:AF31] = ""
    If Mode = 0 Then [AI:AI] = ""
    Cells.Font.ColorIndex = 1
    For R = 2 To 30
        For C = 2 To 27
            If Cells(R, C).Interior.ColorIndex = 6 Then
            Cells(R, C) = Chr(159)
            End If
        Next C
    Next R
' PacMan
    R = 24
    C = 15
    Cells(R, C) = "J"
    Cells(R, C).Font.ColorIndex = 3
' Enemy 1
    R1 = 2
    C1 = 27
    Cells(R1, C1) = "L"
    Cells(R1, C1).Font.ColorIndex = 4
    Z1 = Chr(159)
' Enemy 2
    R2 = 2
    C2 = 2
    Cells(R2, C2) = "L"
    Cells(R2, C2).Font.ColorIndex = 4
    Z2 = Chr(159)
'' Enemy 3
'    R3 = 6
'    C3 = 2
'    Cells(R3, C3) = "L"
'    Cells(R3, C3).Font.ColorIndex = 4
'    Z3 = Chr(159)
'' Enemy  4
'    R4 = 6
'    C4 = 27
'    Cells(R4, C4) = "L"
'    Cells(R4, C4).Font.ColorIndex = 4
'    Z4 = Chr(159)
' Score
    Score = 0
    [AF2] = Score
    [AF1] = "Score"
    Protocol = 1
    If Mode = 1 Then
        Protocol = 1
        Do While Cells(Protocol, 80) <> ""
            If Cells(Protocol, 80) = "U" Then
                UpReplay
            ElseIf Cells(Protocol, 80) = "D" Then
                DownReplay
            ElseIf Cells(Protocol, 80) = "L" Then
                LeftReplay
            ElseIf Cells(Protocol, 80) = "R" Then
                RightReplay
            Else
                MsgBox "Error"
            End If
            Protocol = Protocol + 1
            T = Timer
            Do While Timer - T < 0.4
            Loop
        Loop
    End If
End Sub
Public Sub Start()
    Mode = 0
    If Mode = 0 Then Beep
    Begin
End Sub
Public Sub Replay()
    Mode = 1
    Begin
End Sub

' Replay not working correctly. Need to make record for every enemy. Right now record only for 'Player'

