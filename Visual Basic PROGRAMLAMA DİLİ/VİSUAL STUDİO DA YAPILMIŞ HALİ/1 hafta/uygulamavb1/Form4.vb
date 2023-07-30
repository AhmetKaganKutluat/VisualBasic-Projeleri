Public Class Form4
    Dim ukenar, kkenar, alan, cevre As Integer
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ukenar = TextBox1.Text
        kkenar = TextBox2.Text
        alan = 4 * ukenar
        cevre = 4 * kkenar
        ListBox1.Items.Add("kare alan " & alan)
        ListBox1.Items.Add("kare cevre " & cevre)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ukenar = TextBox1.Text
        kkenar = TextBox2.Text
        alan = ukenar * kkenar
        cevre = ((2 * ukenar) + (2 * kkenar))
        ListBox1.Items.Add("dikdortgen alan " & alan)
        ListBox1.Items.Add("dikdörtgen cevre " & cevre)

    End Sub
End Class