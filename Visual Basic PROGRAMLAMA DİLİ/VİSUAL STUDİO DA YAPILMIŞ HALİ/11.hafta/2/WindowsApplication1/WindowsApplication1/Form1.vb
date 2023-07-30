Public Class Form1
    Dim dizi() As Integer ' boyutsuz bir dizi tanımladık '
    Dim kactane, toplam, ortalama As Integer
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        kactane = Val(InputBox("Kac Tane Sayi Almak Istersiniz ?"))
        '' ReDim dizi(kactane) As Integer
        'daha önceden tanımlanan dizi yeniden yapılandırıldı'
        For i = 1 To kactane
            dizi(i) = Val(InputBox(Str(i) + " . Sayiyi Girin"))
            toplam = toplam + dizi(i)

        Next i
        ortalama = toplam / kactane


        ''  Cls 
        For i = 1 To kactane
            Print("Girilen " + Str(i) + " . sayi = " + Str(dizi(i)))
        Next
        Print("Toplam   = " + Str(toplam))
        Print("Ortalama = " + Str(ortalama))

    End Sub
End Class
