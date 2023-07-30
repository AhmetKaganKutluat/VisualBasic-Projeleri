Public Class Form2
    Dim cik As Integer
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        cik = MsgBox("program kapatılsın mı ", 4 + 32 + 256, "çıkış")
        If (cik = 6) Then

            Application.Exit()

        End If

    End Sub
End Class