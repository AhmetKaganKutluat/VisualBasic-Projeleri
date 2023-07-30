Public Class Form4

    Dim sonek As String
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If RadioButton1.Checked = True Then sonek = ".edu.tr"
        If RadioButton2.Checked = True Then sonek = ".com.tr"
        If RadioButton3.Checked = True Then sonek = ".mil.tr"
        If RadioButton4.Checked = True Then sonek = ".gov.tr"
        If RadioButton5.Checked = True Then sonek = ".org.tr"
        TextBox4.Text = LCase(TextBox1.Text.Trim) & LCase(TextBox2.Text.Trim) & "@" & LCase(TextBox3.Text.Trim) & sonek


    End Sub
End Class