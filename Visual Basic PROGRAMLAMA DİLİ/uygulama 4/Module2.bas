Attribute VB_Name = "Module2"
Sub toplama()
Dim sayi1, sayi2, sonuc As Integer
sayi1 = Val(Form1.Text2.Text)
sayi2 = Val(Form1.Text3.Text)
sonuc = sayi1 + sayi2
MsgBox ("SAYILARIN TOPLAMI  :" & Str(sonuc)) ''str yazdanda olur ama + koymak laz�m yazmasanda
''+ metinsel tipteki ifadelri birle�tirmek i�in ama & her de�erde kullan�l�r
End Sub
