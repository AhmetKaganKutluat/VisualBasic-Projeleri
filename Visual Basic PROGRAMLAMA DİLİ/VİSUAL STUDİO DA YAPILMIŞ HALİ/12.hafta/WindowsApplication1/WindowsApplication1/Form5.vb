Public Class Form5

    Private Sub Form2ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles Form2ToolStripMenuItem.Click
        Form2.MdiParent = Me
        Form2.Show()

    End Sub

    Public Sub formekle(isim As String, renk As String)
        Dim yeniform As New Form
        yeniform.Name = isim
        yeniform.Text = isim
        yeniform.BackColor = Color.FromName(renk)
        yeniform.MdiParent = Me
        yeniform.Show()

    End Sub
    Private Sub OtomatikFormOluşturToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OtomatikFormOluşturToolStripMenuItem.Click
        formekle("Yeni", "Purple")

    End Sub

    Private Sub MehmetToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MehmetToolStripMenuItem.Click
        formekle("Mehmet", "Blue")
    End Sub

    Private Sub AhmetToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AhmetToolStripMenuItem.Click
        formekle("Ahmet", "Green")
    End Sub

    Private Sub FurkanToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FurkanToolStripMenuItem.Click
        formekle("Furkan", "Yellow")
    End Sub
End Class