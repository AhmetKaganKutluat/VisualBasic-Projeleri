Public Class Form2
    Dim isimler(6) As String
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        For i = 1 To 6
            isimler(i) = InputBox(i & " . ismi giriniz ")
        Next i
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        For i = 1 To 6
            ListBox1.Items.Add(isimler(i))
        Next i
    End Sub
End Class