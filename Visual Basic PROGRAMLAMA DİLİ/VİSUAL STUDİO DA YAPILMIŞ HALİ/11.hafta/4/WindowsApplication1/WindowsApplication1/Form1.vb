Public Class Form1
    Dim sayilar(6) As Integer
    Dim negatif_adet, Pozitif_adet, notr_adet As Integer

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        For i = 1 To 6
            sayilar(i) = Val(InputBox(Str(i) + " . sayiyi girin"))
            If (sayilar(i) < 0) Then negatif_adet = negatif_adet + 1
            If (sayilar(i) > 0) Then pozitif_adet = pozitif_adet + 1
            If (sayilar(i) = 0) Then notr_adet = notr_adet + 1
        Next
        'Cls()

        Print("Pozıtıf adet sayısı = " + Str(pozitif_adet))
        Print("Negatif adet sayısı = " + Str(negatif_adet))
        Print("notr adet sayısı = " + Str(notr_adet))
    End Sub
End Class
