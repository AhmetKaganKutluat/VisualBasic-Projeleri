Public Class Form1
    Dim soru As Integer
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '' farklı bir yöntemle dosyaya kayıt işlemi
        If (TextBox1.Text = "") Then
            MsgBox("MÜŞTERİNİN AD BİLGİSİ YAZILSIN")
            TextBox1.Select()
            Exit Sub
        End If

        soru = MsgBox("MÜŞTERİ BİLGİSİ KAYIT YAPILSINMI ", 4 + 32, "KAYIT")

        If (soru = 6) Then '' evetse kayıt yapacak

            My.Computer.FileSystem.WriteAllText("bilgi.txt", TextBox1.Text + Chr(13), True) '' metin belgesine kayıt 

            '' hem kendi dosya oluşturup metni içine yazar
            '' vbcr , chr(13) alt alta yazmak için kullanırlı.
            '' eğer true yerine false dersek bir önceki yazıyı silerek kayıt eder
            '' eğer bunu bir klasörün içinde txt yaratmasını istersek "bilgi.txt" yerine örn = "ogr\bilgi.txt" şeklinde yazılır
            Exit Sub
        End If

        If (soru = 7) Then '' hayırsa dosyadan  okuma yapıcak
            Dim okunan = My.Computer.FileSystem.ReadAllText("bilgi.txt") '' dosyadan okuma komutu 
            RichTextBox1.Text = okunan
            Exit Sub
        End If



    End Sub
End Class
