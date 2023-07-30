Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim i As Integer
        For i = 1 To 3
            ComboBox1.Items.Add(Choose(i, "AKDENİZ BÖLGESİ", "MARMARA BÖLGESİ", "EGE BÖLGESİ"))
        Next
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        ComboBox2.Items.Clear()
        If (ComboBox1.Text = "AKDENİZ BÖLGESİ") Then ''combobox1.selectedindex = 0 şeklindede yazılabilirdi.
            ComboBox2.Items.Add("ANTALYA")
            ComboBox2.Items.Add("MERSİN")
            ComboBox2.Items.Add("ISPARTA")
        ElseIf (ComboBox1.Text = "MARMARA BÖLGESİ") Then
            ComboBox2.Items.Add("İSTANBUL")
            ComboBox2.Items.Add("BURSA")
            ComboBox2.Items.Add("EDİRNE")
        ElseIf (ComboBox1.Text = "EGE BÖLGESİ") Then
            ComboBox2.Items.Add("İZMİR")
            ComboBox2.Items.Add("MANİSA")
            ComboBox2.Items.Add("AYDIN")
        End If
    End Sub

    Dim plaka, sehir, veri As String

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim soru As Integer

        If (TextBox1.Text = "") Then
            MsgBox("boş bırakılamaz")
            Exit Sub
        End If

        veri = TextBox1.Text.Trim   '' kesinlikle trim almak gerekiyor
        plaka = Strings.Left(veri, 2)
        sehir = Strings.Right(veri, Strings.Len(veri) - 2)
        soru = MsgBox(veri & "kaydı yapılsınmı ", 4 + 64, "kayıt işlem")
        If (soru = 6) Then
            ListBox1.Items.Add(plaka)
            ListBox2.Items.Add(sehir)
            TextBox1.Clear()
            Label1.Text = "TOPLAM KAYIT :" + Convert.ToString(ListBox1.Items.Count)  '' yada   & listbox1.items.add

        End If

    End Sub
End Class
