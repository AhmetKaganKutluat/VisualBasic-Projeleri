Public Class Form3

   
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If (ListBox1.SelectedIndex < 0) Then
            MsgBox("eklenecek ders adı seçilmedi lütfen seçiniz")
            Exit Sub
        End If

        ListBox2.Items.Add(ListBox1.Text)
        ListBox1.Items.Remove(ListBox1.SelectedIndex)
    End Sub

    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ListBox1.Items.Add("NESNE TABANLI PROGRAMLA")
        ListBox1.Items.Add("GÖRSEL  PROGRAMLA")
        ListBox1.Items.Add("MAT")
        ListBox1.Items.Add("TARİH")
        ListBox1.Items.Add("BEDEN")
    End Sub

   
    
    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        If (ListBox1.Items.Count = 0) Then
            Button1.Enabled = False
        End If
    End Sub

    Private Sub ListBox1_Click(sender As Object, e As EventArgs) Handles ListBox1.Click
        Label1.Text = ("DURUM" & ListBox1.SelectedIndex + 1 & " / " & ListBox1.Items.Count)
    End Sub
End Class