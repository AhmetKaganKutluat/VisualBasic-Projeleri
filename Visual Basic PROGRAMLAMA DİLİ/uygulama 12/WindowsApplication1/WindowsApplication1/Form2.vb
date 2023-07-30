Public Class Form2

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim ad, soyad As String
        ad = TextBox1.Text.Trim()
        soyad = TextBox2.Text.Trim()
        Form3.TextBox1.Text = ad
        Form3.TextBox2.Text = soyad
        Form3.ShowDialog()




    End Sub
End Class