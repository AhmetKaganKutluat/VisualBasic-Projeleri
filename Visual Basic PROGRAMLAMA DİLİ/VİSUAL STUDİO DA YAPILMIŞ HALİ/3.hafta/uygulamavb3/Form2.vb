Public Class Form2
    Dim ortalama, topla, sayilar(5), i As Integer
    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        For i = 1 To 5
            sayilar(i) = Val(InputBox(i & " . Sayiyi giriniz ", "SAYI GİRİŞİ"))
            topla = topla + sayilar(i)
        Next i
        ortalama = topla / 5
        Label1.Text = ("TOPLAM : " & topla & " ORTALAMA  " & ortalama)
    End Sub
End Class