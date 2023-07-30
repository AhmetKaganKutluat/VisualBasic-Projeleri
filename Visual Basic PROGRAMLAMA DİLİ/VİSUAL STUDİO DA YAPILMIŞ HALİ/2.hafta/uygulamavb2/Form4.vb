Public Class Form4

    Dim isim As String
    Dim sayi As Integer
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        sayi = TextBox1.Text
        isim = TextBox2.Text
        If (sayi < 50) Then
            Label1.Text = isim & " " & "SINIFTA KALDI"
        Else
            Label1.Text = isim & " " & "SINIFI GEÇTİ"
        End If
    End Sub

End Class
