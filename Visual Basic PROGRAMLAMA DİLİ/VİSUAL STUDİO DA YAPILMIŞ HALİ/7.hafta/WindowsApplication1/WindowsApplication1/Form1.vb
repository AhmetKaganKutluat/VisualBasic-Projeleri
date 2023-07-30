Public Class Form1
    Public durum1 As Boolean
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        durum1 = True
        ComboBox1.Items.Add("MARMARA BÖLGESİ")
        ComboBox1.Items.Add("AKDENİZ BÖLGESİ")
        ComboBox1.Items.Add("EGE BÖLGESİ")
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If (ComboBox1.SelectedIndex = 0) Then ''listındex yerine "MARMARA BÖLGESİ DE YAZILABİLİRDİ"
            ComboBox2.Items.Clear()
            ComboBox2.Items.Add("İSTANBUL")
            ComboBox2.Items.Add("EDİRNE")
            ComboBox2.Items.Add("KOCAELİ")
            ComboBox2.Items.Add("BURSA")
        End If

        If (ComboBox1.SelectedIndex = 1) Then ''listındex yerine "AKDENİZ BÖLGESİ DE YAZILABİLİRDİ"
            ComboBox2.Items.Clear()
            ComboBox2.Items.Add("MERSİN")
            ComboBox2.Items.Add("ADANA")
            ComboBox2.Items.Add("ANTALYA")
            ComboBox2.Items.Add("BURDUR")
        End If

        If (ComboBox1.SelectedIndex = 2) Then ''listındex yerine "EGE BÖLGESİ DE YAZILABİLİRDİ"
            ComboBox2.Items.Clear()

            ComboBox2.Items.Add("İZMİR")
            ComboBox2.Items.Add("MANİSA")
            ComboBox2.Items.Add("AYDIN")
            ComboBox2.Items.Add("MUĞLA")
        End If
    End Sub
End Class
