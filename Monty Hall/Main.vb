Public Class Main

    Declare Sub Sleep Lib "kernel32.dll" (ByVal Milliseconds As Integer)

    Private Switched As Boolean

    Private Delay As Integer = 100000

    Private Sub PlayGame()
        While vbTrue
            If 100000 = Delay Then
                Exit Sub
            End If

            Announce.Text = "Hiding Car"

            Me.Refresh()

            Pick1.Checked = False
            Pick2.Checked = False
            Pick3.Checked = False

            Door1.BackColor = Drawing.Color.LightGray
            Door2.BackColor = Drawing.Color.LightGray
            Door3.BackColor = Drawing.Color.LightGray

            'Hide Prize

            Select Case CInt(Math.Floor((3 - 1 + 1) * Rnd())) + 1
                Case 1
                    Door1.Text = "Car"
                    Door2.Text = "Goat"
                    Door3.Text = "Goat"
                Case 2
                    Door1.Text = "Goat"
                    Door2.Text = "Car"
                    Door3.Text = "Goat"
                Case Else
                    Door1.Text = "Goat"
                    Door2.Text = "Goat"
                    Door3.Text = "Car"
            End Select

            Me.Refresh()

            If 0 <> Delay Then
                System.Threading.Thread.Sleep(Delay)
            End If

            'Contestant's choice

            Announce.Text = "Select Door"

            Select Case CInt(Math.Floor((3 - 1 + 1) * Rnd())) + 1
                Case 1
                    Pick1.Checked = True
                    Pick2.Checked = False
                    Pick3.Checked = False
                Case 2
                    Pick1.Checked = False
                    Pick2.Checked = True
                    Pick3.Checked = False
                Case Else
                    Pick1.Checked = False
                    Pick2.Checked = False
                    Pick3.Checked = True
            End Select

            Me.Refresh()

            If 0 <> Delay Then
                System.Threading.Thread.Sleep(Delay)
            End If

            'Monty's choice

            Announce.Text = "You didn't pick..."

            If True = Pick1.Checked Then
                If "Car" = Door2.Text Then
                    Door3.BackColor = Drawing.Color.White
                Else
                    Door2.BackColor = Drawing.Color.White
                End If
            ElseIf True = Pick2.Checked Then
                If "Car" = Door1.Text Then
                    Door3.BackColor = Drawing.Color.White
                Else
                    Door1.BackColor = Drawing.Color.White
                End If
            Else
                If "Car" = Door1.Text Then
                    Door2.BackColor = Drawing.Color.White
                Else
                    Door1.BackColor = Drawing.Color.White
                End If
            End If

            Me.Refresh()

            If 0 <> Delay Then
                System.Threading.Thread.Sleep(Delay)
            End If

            'Player switch

            Announce.Text = "Wanna switch?"

            If 1 = CInt(Math.Floor((2 - 1 + 1) * Rnd())) + 1 Then
                If True = Pick1.Checked Then
                    Pick1.Checked = False

                    If Drawing.Color.LightGray <> Door2.BackColor Then
                        Pick3.Checked = True
                    Else
                        Pick2.Checked = True
                    End If
                ElseIf True = Pick2.Checked Then
                    Pick2.Checked = False

                    If Drawing.Color.LightGray <> Door1.BackColor Then
                        Pick3.Checked = True
                    Else
                        Pick1.Checked = True
                    End If
                Else
                    Pick3.Checked = False

                    If Drawing.Color.LightGray <> Door2.BackColor Then
                        Pick1.Checked = True
                    Else
                        Pick2.Checked = True
                    End If
                End If
                Switched = True
            Else
                Switched = False
            End If

            Me.Refresh()

            If 0 <> Delay Then
                System.Threading.Thread.Sleep(Delay)
            End If

            'Win or lose

            If (Pick1.Checked And "Car" = Door1.Text) Or (Pick2.Checked And "Car" = Door2.Text) Or (Pick3.Checked And "Car" = Door3.Text) Then
                Announce.Text = "You Win"

                If Switched Then
                    SwitchWins.Text = Str(Val(SwitchWins.Text) + 1)
                Else
                    StayWins.Text = Str(Val(StayWins.Text) + 1)
                End If
            Else
                Announce.Text = "You Lose"

                If Switched Then
                    SwitchLosses.Text = Str(Val(SwitchLosses.Text) + 1)
                Else
                    StayLosses.Text = Str(Val(StayLosses.Text) + 1)
                End If
            End If

            Me.Refresh()

            If 0 <> Delay Then
                System.Threading.Thread.Sleep(Delay)
            End If

            System.Windows.Forms.Application.DoEvents()
        End While
    End Sub

    Private Sub Pause_Click(sender As System.Object, e As System.EventArgs) Handles Pause.Click
        Delay = 100000
    End Sub

    Private Sub Slow_Click(sender As System.Object, e As System.EventArgs) Handles Slow.Click
        If 100000 = Delay Then
            Delay = 2000
            PlayGame()
        Else
            Delay = 2000
        End If
    End Sub

    Private Sub Fast_Click(sender As System.Object, e As System.EventArgs) Handles Fast.Click
        If 100000 = Delay Then
            Delay = 500
            PlayGame()
        Else
            Delay = 500
        End If
    End Sub

    Private Sub Full_Click(sender As System.Object, e As System.EventArgs) Handles Full.Click
        If 100000 = Delay Then
            Delay = 0
            PlayGame()
        Else
            Delay = 0
        End If
    End Sub

    Private Sub Main_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Randomize()
    End Sub

End Class