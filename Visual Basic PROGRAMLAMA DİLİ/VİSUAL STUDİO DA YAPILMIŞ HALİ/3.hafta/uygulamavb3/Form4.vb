Public Class Form4

   
    Sub hesap(sayi) ''procedure
        fibo(1) = 1
        For i = 2 To sayi
            fibo(i) = fibo(i - 1) + fibo(i - 2)
        Next i
        For i = 1 To sayi
            ListBox1.Items.Add(i & ".  fibo sayısı : " & fibo(i))
        Next i
    End Sub
    Dim fibo(500), sayi As Integer
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ListBox1.Items.Clear()
        sayi = Val(InputBox("KAÇ TANE FİBO SAYISI GÖSTERİLSİN ", "FİBO SERİSİ"))
        hesap(sayi)
    End Sub
End Class