Public Class Form3

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        If (DateTimePicker1.Value > DateTimePicker2.Value) Then
            MsgBox("ilk grilen tarih ikinci tarihten büyük olamaz")

        End If
    End Sub
    Dim t1, t2, sonuc As Integer
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        t1 = Val(DateTimePicker1.Value)
        t2 = Val(DateTimePicker2.Value)
        sonuc = t2 - t1
        sonuc = sonuc * -1
        TextBox1.Text = sonuc
        MsgBox(DateTimePicker1.Text)

    End Sub
End Class