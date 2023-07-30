Public Class Form2

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        ProgressBar1.Value += 2
        Label1.Text = "program yükleniyor bekleyiniz  = % " & ProgressBar1.Value
        If (ProgressBar1.Value = 100) Then
            Timer1.Enabled = False
            Form3.ShowDialog()
            Me.Close()
        End If

    End Sub
End Class