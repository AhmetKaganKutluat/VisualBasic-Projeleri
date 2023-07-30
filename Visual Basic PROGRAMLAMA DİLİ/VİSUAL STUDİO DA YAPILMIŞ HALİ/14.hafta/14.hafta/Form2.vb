Public Class Form2

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ''RichTextBox1.Text = (My.Computer.FileSystem.ReadAllText("bölge.txt"))
        ''For i = 0 To RichTextBox1.Lines.Count - 1
        ''    ComboBox1.Items.Add(RichTextBox1.Text)
        ''  Next
        Dim okunan As String
        okunan = My.Computer.FileSystem.ReadAllText("bölge.txt")

        Dim bol() As String
        bol = Split(okunan, Environment.NewLine)   '' alt alta yazmak için enter basılan yerden bolmem gerek
        For Each veriler As String In bol
            ComboBox1.Items.Add(veriler)
        Next
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Process.Start("www.akdeniz.edu.tr")
    End Sub
End Class