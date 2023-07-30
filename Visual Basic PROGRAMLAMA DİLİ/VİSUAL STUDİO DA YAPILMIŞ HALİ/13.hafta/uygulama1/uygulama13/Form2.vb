Public Class Form2
    Dim s1, s2 As Integer
    Dim sonuc As Double
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        s1 = TextBox1.Text
        s2 = TextBox2.Text
        If (RadioButton1.Checked = True) Then               ''elseif olarakda yapılabilirdi..
            sonuc = s1 + s2
            TextBox3.Text = "işlem sonucu : " & sonuc
            Exit Sub
        End If
        If (RadioButton2.Checked = True) Then
            sonuc = s1 - s2
            TextBox3.Text = "işlem sonucu : " & sonuc
            Exit Sub
        End If
        If (RadioButton3.Checked = True) Then
            sonuc = s1 * s2
            TextBox3.Text = "işlem sonucu : " & sonuc
            Exit Sub
        End If
        If (RadioButton4.Checked = True) And (s2 <> 0) Then
            sonuc = s1 / s2
            TextBox3.Text = "işlem sonucu : " & sonuc
            Exit Sub
        End If

    End Sub
End Class