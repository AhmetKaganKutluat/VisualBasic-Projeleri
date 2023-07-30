Public Class Form2

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        Form1.Show()
        If (Form1.durum1 = True) Then
            ToolStripButton2.Enabled = False
        End If
    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        Form3.Show()
        If (Form3.durum = True) Then
            ToolStripButton2.Enabled = False
        End If
    End Sub

    
    '' Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
    '   If (Form3.durum = False) Then
    ''     ToolStripButton2.Enabled = True
    '' End If
    'If (Form1.durum1 = False) Then
    '   ToolStripButton2.Enabled = True
    'End If
    'End Sub
End Class