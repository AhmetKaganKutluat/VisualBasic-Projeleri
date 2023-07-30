Public Class Form1
    Dim dizi(2, 5) As Integer
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        On Error Resume Next 'hata varsa sıradankıne gec'
        For i = 1 To 5
            dizi(1, i) = InputBox(Str(i) + " . elemanı girin ")
            dizi(2, i) = dizi(1, i) * dizi(1, i)
        Next

        For i = 1 To 5
            ''Print Str(dizi(1, i)) + " sayının karesi = " + Str(dizi(2, i))
            ListBox1.Items.Add(Str(dizi(1, i)) + " Sayının karesi = " + Str(dizi(2, i)))
        Next
    End Sub
End Class
