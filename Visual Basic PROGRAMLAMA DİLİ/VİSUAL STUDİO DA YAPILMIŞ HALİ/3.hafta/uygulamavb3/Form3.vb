Public Class Form3
    Function hesapla(a, b)

        alan = a * b
        cevre = 2 * (a + b)
        hesapla = ("ALAN: " & alan & " ÇEVRE :" & cevre)
    End Function
    Dim a, b, alan, cevre, sayi, us As Integer

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        a = Convert.ToInt16(TextBox1.Text)
        b = Convert.ToInt16(TextBox2.Text)
        Label1.Text = hesapla(a, b)
        TextBox1.Text = ""
        TextBox2.Text = ""
    End Sub
    Function kuvvet(sayi, us)
        kuvvet = 1
        For i = 1 To us
            kuvvet = kuvvet * sayi
        Next i
    End Function
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        On Error Resume Next
        sayi = InputBox("TABAN SAYİSİNİ GİRİNİZ  :", "SAYI GİRİŞİ")
        us = InputBox("ÜST KUVVETİNİ GİRİNİZ  :", "SAYI GİRİŞİ")
        MsgBox("SAYININ KUVVETİ  : " & kuvvet(sayi, us))
    End Sub
End Class