Public Class Form1
    Dim a, b, c, delta, tekkok, x1, x2 As Integer
    Dim parabol As String
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        a = Val(TextBox1.Text)
        b = Val(TextBox2.Text)
        c = Val(TextBox3.Text)

        delta = (b * b) - (4 * a * c)
        Label6.Text = "DELTA SONUCU :" & delta

        If (a < 0) Then
            parabol = "AŞAĞI YÖNLÜ PARABOL "
        Else
            parabol = "YUKARI YÖNLÜ PARABOL"
        End If

        Label4.Text = parabol

        If (delta < 0) Then
            Label5.Text = "FONKSİYONUN REEL KÖKÜ YOKTUR."
        Else
            If (delta = 0) Then
                tekkok = (-1 * b) / (2 * a)
                Label5.Text = "TEK KÖK VARDIR . KÖK DEĞERİ :" & tekkok
            Else
                If (delta > 0) Then
                    x1 = ((-1 * b) + Math.Sqrt(delta)) / (2 * a) '' math.sgr karakök demek sgr karakök demektir.  (c# da sgrt oluyor t harfi ekleniyor)
                    x2 = ((-1 * b) - Math.Sqrt(delta)) / (2 * a)
                    x1 = Math.Round(x1, 2) '' iki basamak yuvarladı
                    x2 = Math.Round(x2, 2)
                    Label5.Text = ("ÇİFT KÖK VARDIR X1 : " & x1 & " X2  :" & x2)
                End If
            End If
        End If



    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If Trim(TextBox1.Text) = "" Or Trim(TextBox2.Text) = "" Or Trim(TextBox3.Text) = "" Then
            Button1.Enabled = False
        Else
            Button1.Enabled = True
        End If
    End Sub
End Class
