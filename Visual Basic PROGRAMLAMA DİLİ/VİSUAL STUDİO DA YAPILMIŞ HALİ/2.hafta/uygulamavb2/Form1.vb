Public Class Form1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If (Trim(TextBox1.Text) = "") Then
            MsgBox("LÜTFEN MÜŞTERİNİN ADINI YAZINIZ" & Chr(13) & "                   HATA YAPTINIZ")
        Else
            If (Trim(TextBox2.Text) = "") Then
                MsgBox("LÜTFEN MÜŞTERİNİN SOYADINI YAZINIZ" & Chr(13) & "                   HATA YAPTINIZ")
            Else
                If (ComboBox1.Text = "") Then
                    MsgBox("lütfen bir isim seçiniz")
                Else
                    MsgBox(" BAŞARILI ")
                End If
            End If
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.Items.Add("beden eğitimi")
        ComboBox1.Items.Add("görsel programlama")
        ComboBox1.Items.Add("matematik")
        ComboBox1.Items.Add("nesne tabanlı programlama ")
        ComboBox1.Items.Add("tarih")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Form2.ShowDialog()
    End Sub
End Class
