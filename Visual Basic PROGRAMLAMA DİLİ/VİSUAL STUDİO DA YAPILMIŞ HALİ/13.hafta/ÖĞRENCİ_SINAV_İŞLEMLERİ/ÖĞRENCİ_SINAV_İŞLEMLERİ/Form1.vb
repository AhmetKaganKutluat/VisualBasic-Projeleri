Public Class Form1

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged

        If (CheckBox1.Checked = True) Then           '' göster gizle komutu için 
            TextBox2.PasswordChar = ""
        Else
            TextBox2.PasswordChar = "*"
        End If



    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        If (e.KeyChar = Chr(13)) Then
            Button1.Select()
        End If
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If (e.KeyChar = Chr(13)) Then
            TextBox2.Select()
        End If

    End Sub

    Dim hak As Integer = 3  '' hak 3 değeri atadık ve bu hakkı asla butona tanımlamamak gerek yoksa hep yeniler kendini

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim kullanıcı, sifre As String

        kullanıcı = TextBox1.Text.Trim
        sifre = TextBox2.Text.Trim
        If (kullanıcı = "furkan") And (sifre = "2020") Then
            Form3.ToolStripStatusLabel1.Text = "aktif kullanıcı : " & kullanıcı

            Me.Hide()
            Form2.ShowDialog()
            '' Application.Exit()da diyebilirdik
            Me.Close()

        Else

            hak -= 1
            MsgBox("ŞİFRE YADA KULLANCI ADI HATALI " & vbCr & "kalan hakkınız :" & hak)
            If (hak = 0) Then
                Me.Close()

            End If

            TextBox1.Clear()
            TextBox2.Clear()
            Exit Sub
        End If
    End Sub

   
  
End Class
