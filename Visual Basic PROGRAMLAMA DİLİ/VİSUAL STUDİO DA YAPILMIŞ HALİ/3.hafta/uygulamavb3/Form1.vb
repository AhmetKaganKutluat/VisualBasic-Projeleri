Public Class Form1

    Dim secim As Object
    Dim ad, soyad, adres As String
    Dim gelir, gider, fark As Integer

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If (TextBox1.Text = "") Then
            MsgBox("lütfen adı kısmını doldurunuz .")
            Exit Sub
        End If
        If (TextBox2.Text = "") Then
            MsgBox("lütfen soyadı kısmını doldurunuz .")
            Exit Sub
        End If
        If (TextBox3.Text = "") Then
            MsgBox("gelir kısmını doldurunuz.")
            Exit Sub
        End If
        If (TextBox4.Text = "") Then
            MsgBox("gider kısmını doldurunuz.")
            Exit Sub
        End If
        If (TextBox5.Text = "") Then
            MsgBox("lütfen adres kısmını doldurunuz .")
            Exit Sub
        End If
git:
        secim = InputBox("LÜTFEN AŞAĞIDAKİ SEÇİMLERDEN BİRİNİ YAPINIZ", "Lütfen Seçiniz", " 1 - 9 , HESAPLA ,5,S ")
        secim = UCase(Trim(secim)) ''büyük harfe çevirmek için

        Select Case secim
            Case 1 To 9 : MsgBox("görsel programlama")
                GoTo git
            Case "S" : End
            Case "5" : MsgBox("SELAM KACMAZ " & vbCr & "MANAVGAT MYO")
            Case "HESAPLA"
                ad = TextBox1.Text.Trim
                soyad = TextBox2.Text.Trim
                gelir = Convert.ToInt16(TextBox3)  ''cint ordaki değeri sayısala dök ifadesidir.
                gider = Convert.ToInt16(TextBox4)
                adres = TextBox5.Text.Trim
                fark = gelir - gider
                fark = fark * -1
                MsgBox(ad & " " & soyad & fark & " tl maaşlı " & vbCr & adres & "adresinde yaşamaktadır.")
                GoTo git
            Case Else
                MsgBox("tanımsız  tuşlama  ")
                TextBox1.Text = ""
                TextBox2.Text = ""
                TextBox3.Text = ""
                TextBox4.Text = ""
                TextBox5.Text = ""
        End Select
    End Sub
End Class
