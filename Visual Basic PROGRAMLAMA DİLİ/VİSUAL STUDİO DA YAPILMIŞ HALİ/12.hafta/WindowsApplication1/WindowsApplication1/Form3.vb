Public Class Form3

    Private Sub Form3_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        Form2.TextBox1.Clear()
        Form2.TextBox2.Clear()
        Form2.TextBox1.BackColor = Color.DarkGreen
        Form2.TextBox2.BackColor = Color.DarkMagenta

    End Sub

    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load



    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Me.Text = "Ahmet Furkan" & DateTime.Now
    End Sub
End Class