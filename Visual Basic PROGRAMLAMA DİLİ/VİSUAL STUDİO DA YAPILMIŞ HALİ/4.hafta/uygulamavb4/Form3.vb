Public Class Form3
    Function turet()
        Dim sifre As String
        Dim karakterler(25)
        karakterler(0) = "A"
        karakterler(1) = "B"
        karakterler(2) = "T"
        karakterler(3) = "6"
        karakterler(4) = "8"
        karakterler(5) = "{"
        karakterler(6) = "S"
        karakterler(7) = "F"
        karakterler(8) = "7"
        karakterler(9) = "+"
        karakterler(10) = "P"
        karakterler(11) = "?"
        karakterler(12) = "M"
        karakterler(13) = "m"
        karakterler(14) = "r"
        karakterler(15) = "H"
        karakterler(16) = "w"
        karakterler(17) = "X"
        karakterler(18) = "/"
        karakterler(19) = "*"
        karakterler(20) = "#"
        karakterler(21) = "&"
        karakterler(22) = "t"
        karakterler(23) = "7"
        karakterler(24) = "j"
        karakterler(25) = "H"
        For i = 1 To 8 '' 8 basamalklı şifre
            Randomize()
            sifre = sifre & karakterler(Int(Rnd() * 25))
        Next i
        turet = sifre
    End Function
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TextBox1.Text = turet
    End Sub
End Class