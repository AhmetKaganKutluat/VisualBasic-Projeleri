Public Class Form1
    Dim sayi(5) As Integer
    Dim gecici As Integer
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        For i = 1 To 5
            sayi(i) = Val(InputBox(Str(i) + " . sayıyı girin"))


        Next
        For i = 1 To 5
            For j = 1 To 5

                If (sayi(j) > sayi(i)) Then
                    gecici = sayi(i)

                    ''5       8      7     3      6

                    sayi(i) = sayi(j)
                    sayi(j) = gecici

                End If
            Next
        Next
        ListBox1.Items.Clear()


        For i = 1 To 5
            ListBox1.Items.Add(sayi(i))
        Next
    End Sub
End Class
