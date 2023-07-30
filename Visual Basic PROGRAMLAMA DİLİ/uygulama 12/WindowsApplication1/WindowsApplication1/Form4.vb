Public Class Form4
    Dim sayac As Integer
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If (TextBox1.Text = "" Or TextBox2.Text = "") Then
            MsgBox("KULLANICI ADI VE ŞİFRE BOŞ BIRAKILAMAZ")
            Exit Sub
        End If
        If (TextBox1.Text = "AHMET" And TextBox2.Text = "FURKAN") Then

            Form5.ShowDialog()
            Me.Close()
        Else
            sayac -= 1
            MsgBox("KULLANICI ADI VEYA ŞİFRE YANLIS.." & vbCr & vbCr & _
                   "Kalan Hakkınız =" & sayac) 'vbcr enter demek   alt cizgide alttan devam etmek için'
            If (sayac = 0) Then
                Me.Close()

            End If
        End If
    End Sub

    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        sayac = 3
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If (e.KeyChar = Chr(13)) Then
            TextBox2.Select()
        End If
    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        If (e.KeyChar = Chr(13)) Then
            Button1.Select() 'focus yerine kullanılır.'
        End If
    End Sub
End Class