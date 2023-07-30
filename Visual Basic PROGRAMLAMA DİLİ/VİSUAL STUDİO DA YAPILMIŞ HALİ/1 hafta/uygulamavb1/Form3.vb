Public Class Form3
    Dim s1, s2, sonuc As Integer
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        s1 = CInt(TextBox1.Text)
        s2 = CInt(TextBox2.Text)
        sonuc = s1 + s2
        Label1.Text = "sayıların toplamı :" & sonuc
    End Sub
End Class