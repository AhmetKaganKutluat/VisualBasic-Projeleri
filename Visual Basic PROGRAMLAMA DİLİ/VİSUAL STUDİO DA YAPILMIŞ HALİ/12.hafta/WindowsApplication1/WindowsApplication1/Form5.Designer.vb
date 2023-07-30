<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form5
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.DosyaToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Form2ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripSeparator()
        Me.OtomatikFormOluşturToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.YeniToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MehmetToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AhmetToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.FurkanToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.DosyaToolStripMenuItem, Me.YeniToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(990, 24)
        Me.MenuStrip1.TabIndex = 1
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'DosyaToolStripMenuItem
        '
        Me.DosyaToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.Form2ToolStripMenuItem, Me.ToolStripMenuItem1, Me.OtomatikFormOluşturToolStripMenuItem})
        Me.DosyaToolStripMenuItem.Name = "DosyaToolStripMenuItem"
        Me.DosyaToolStripMenuItem.Size = New System.Drawing.Size(51, 20)
        Me.DosyaToolStripMenuItem.Text = "Dosya"
        '
        'Form2ToolStripMenuItem
        '
        Me.Form2ToolStripMenuItem.Name = "Form2ToolStripMenuItem"
        Me.Form2ToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.Form2ToolStripMenuItem.Text = "Form2"
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(194, 6)
        '
        'OtomatikFormOluşturToolStripMenuItem
        '
        Me.OtomatikFormOluşturToolStripMenuItem.Name = "OtomatikFormOluşturToolStripMenuItem"
        Me.OtomatikFormOluşturToolStripMenuItem.Size = New System.Drawing.Size(197, 22)
        Me.OtomatikFormOluşturToolStripMenuItem.Text = "Otomatik Form Oluştur"
        '
        'YeniToolStripMenuItem
        '
        Me.YeniToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MehmetToolStripMenuItem, Me.AhmetToolStripMenuItem, Me.FurkanToolStripMenuItem})
        Me.YeniToolStripMenuItem.Name = "YeniToolStripMenuItem"
        Me.YeniToolStripMenuItem.Size = New System.Drawing.Size(41, 20)
        Me.YeniToolStripMenuItem.Text = "Yeni"
        '
        'MehmetToolStripMenuItem
        '
        Me.MehmetToolStripMenuItem.Name = "MehmetToolStripMenuItem"
        Me.MehmetToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.MehmetToolStripMenuItem.Text = "mehmet"
        '
        'AhmetToolStripMenuItem
        '
        Me.AhmetToolStripMenuItem.Name = "AhmetToolStripMenuItem"
        Me.AhmetToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.AhmetToolStripMenuItem.Text = "ahmet"
        '
        'FurkanToolStripMenuItem
        '
        Me.FurkanToolStripMenuItem.Name = "FurkanToolStripMenuItem"
        Me.FurkanToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.FurkanToolStripMenuItem.Text = "furkan"
        '
        'Form5
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(990, 471)
        Me.Controls.Add(Me.MenuStrip1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.IsMdiContainer = True
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "Form5"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "MÜŞTERİ İŞLEMLERİ"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents DosyaToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Form2ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents OtomatikFormOluşturToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents YeniToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MehmetToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AhmetToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents FurkanToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
End Class
