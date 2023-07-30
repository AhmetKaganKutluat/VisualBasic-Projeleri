Public Class Form1
    Dim metin As String
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        metin = Trim(UCase(TextBox1.Text))
        Form2.Show()
        Form2.Label1.Text = metin


    End Sub
End Class
