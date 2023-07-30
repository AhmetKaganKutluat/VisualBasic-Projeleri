Module Module2
    Sub toplama()
        Dim sayi1, sayi2, sonuc As Integer
        sayi1 = Convert.ToInt16(Form1.TextBox2.Text)
        sayi2 = Convert.ToInt16(Form1.TextBox3.Text)
        sonuc = sayi1 + sayi2
        MsgBox("SAYILARIN TOPLAMI  :" & Str(sonuc)) 
    End Sub
End Module
