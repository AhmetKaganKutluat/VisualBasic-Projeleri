Public Class Form1
    Dim soru As Integer
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        soru = MsgBox("Çikis Yapilsin Mi ?", 4 + 32 + 256, "Çikis İslemi")
        If (soru = 6) Then
            ''end
            Application.Exit()

        End If




    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ''ListBox1.Items.Add("GALSAM OTU")
        ''ListBox1.Items.Add("ADAM OTU")
        ''ListBox1.Items.Add("MÜRVER ASA")
        ''ListBox1.Items.Add("GÖRÜNMEZLİK PELERİNİ")
        ''ListBox1.Items.Add("FELSEFE TAŞI")
        Dim i As Integer
        For i = 1 To 5
            ListBox1.Items.Add(Choose(i, "GALSAM OTU", "ADAM OTU", "MÜRVER ASA", "GÖRÜNMEZLİK PELERİNİ", "FELSEFE TAŞI"))
        Next





    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If (ListBox1.SelectedIndex < 0) Then
            MsgBox("Sol Taraftan Seçim Yap Ulan !!!")
            Exit Sub
        End If
        ListBox2.Items.Add(ListBox1.Text)
        ListBox1.Items.RemoveAt(ListBox1.SelectedIndex)


    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If (ListBox1.Items.Count < 1) Then
            Button2.Enabled = False
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        For i = 0 To ListBox1.Items.Count - 1
            ListBox2.Items.Add(ListBox1.Items(i))
        Next
        ListBox1.Items.Clear()

    End Sub
End Class
