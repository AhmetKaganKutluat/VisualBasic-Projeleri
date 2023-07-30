Public Class Form1
    Dim soru As Integer

    Dim adi, soyadi, adres As String
    Sub temizle()
        TextBox1.Text = ""
        TextBox2.Text = ""
        RichTextBox1.Text = ""

    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If (TextBox1.Text = "") Then
            MsgBox("MÜŞTERİNİN ADINI YAZMADINIZ LÜTFEN YAZINIZ ")
            Exit Sub
        End If
        If (TextBox2.Text = "") Then
            MsgBox("MÜŞTERİNİN SOYADINI YAZMADINIZ LÜTFEN YAZINIZ ")
            Exit Sub
        End If
        If (RichTextBox1.Text = "") Then
            MsgBox("MÜŞTERİNİN ADRESİNİ YAZMADINIZ LÜTFEN YAZINIZ ")
            Exit Sub
        End If

        soru = MsgBox(UCase(Trim(TextBox1.Text)) + " " + UCase(Trim(TextBox2.Text)) + " isimi kayıt yapılsınmı ?  ", 4 + 32 + 256, "sisteme kayıt ")

        If (soru = 6) Then '' dosya oluşturup kayıt yapma
            '  Open "kayit.txt" For Append As #1
            adi = Trim(UCase(TextBox1.Text))
            soyadi = Trim(UCase(TextBox2.Text))
            adres = Trim(UCase(RichTextBox1.Text))
            Print(1, adi, soyadi, adres)
            ' Close #1
            MsgBox("bilgiler sisteme kayıt yapılmıştır")
            temizle()
        Else
            temizle() ''procedure

        End If '' if sonu
    End Sub
End Class
