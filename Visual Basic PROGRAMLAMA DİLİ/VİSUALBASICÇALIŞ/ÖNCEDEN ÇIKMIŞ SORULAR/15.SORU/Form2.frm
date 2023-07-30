VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   6360
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15540
   LinkTopic       =   "Form2"
   ScaleHeight     =   6360
   ScaleWidth      =   15540
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   5715
      Left            =   8160
      TabIndex        =   24
      Top             =   240
      Width           =   7095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Hesapla"
      Height          =   735
      Left            =   840
      TabIndex        =   23
      Top             =   3720
      Width           =   1935
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   840
      TabIndex        =   22
      Top             =   3120
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   840
      TabIndex        =   21
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   840
      TabIndex        =   20
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CheckBox Check3 
      Caption         =   "2.Sýnýf"
      Height          =   495
      Left            =   7320
      TabIndex        =   16
      Top             =   1680
      Width           =   735
   End
   Begin VB.CheckBox Check2 
      Caption         =   "1.Sýnýf"
      Height          =   375
      Left            =   6480
      TabIndex        =   15
      Top             =   1800
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Tekrar"
      Height          =   315
      Left            =   5520
      TabIndex        =   14
      Top             =   1800
      Width           =   855
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   5520
      TabIndex        =   12
      Top             =   1320
      Width           =   1935
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   5520
      TabIndex        =   10
      Top             =   840
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   5520
      TabIndex        =   8
      Top             =   360
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   840
      MaxLength       =   15
      TabIndex        =   6
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   840
      MaxLength       =   15
      TabIndex        =   5
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ID Yenile"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label10 
      Caption         =   "Ödev ="
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label Label9 
      Caption         =   "Final ="
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "Vize ="
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label7 
      Caption         =   "Sýnýf ="
      Height          =   255
      Left            =   4920
      TabIndex        =   13
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "Dersler ="
      Height          =   255
      Left            =   4680
      TabIndex        =   11
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Öðretim Elemaný ="
      Height          =   255
      Left            =   4080
      TabIndex        =   9
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Bölüm ="
      Height          =   255
      Left            =   4800
      TabIndex        =   7
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Soyadý ="
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Adý="
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "ID ="
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   495
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vize, final, ödev, evize, efinal, eödev, ortalama As Double
Dim ID, ad, soyad, ders, hoca, bölüm, durum, sinif, harfnotu As String


Function turet()
Dim sifre As String ''text kutusuna aktarýlan þifremdir
Dim karakterler(25) ''kaç tane karakter varsa o kdr ekle
karakterler(0) = "A"
karakterler(1) = "B"
karakterler(2) = "T"
karakterler(3) = "6"
karakterler(4) = "8"
karakterler(5) = "{"
karakterler(6) = "S"
karakterler(7) = "F"
karakterler(8) = "7"
karakterler(9) = "+"
karakterler(10) = "P"
karakterler(11) = "?"
karakterler(12) = "M"
karakterler(13) = "m"
karakterler(14) = "r"
karakterler(15) = "H"
karakterler(16) = "w"
karakterler(17) = "X"
karakterler(18) = "/"
karakterler(19) = "*"
karakterler(20) = "#"
karakterler(21) = "&"
karakterler(22) = "t"
karakterler(23) = "7"
karakterler(24) = "j"
karakterler(25) = "H"
For i = 1 To 5 '' 8 basamalklý þifre
Randomize
sifre = sifre & karakterler(Int(Rnd() * 25))
Next i
turet = sifre
End Function

Sub temizle()
Text1.Text = turet()
List1.Clear
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Combo1.Clear
Combo2.Clear
Combo3.Clear
Check1.Value = False
Check2.Value = False
Check3.Value = False

End Sub




Private Sub Check1_Click()
If (Check1.Value = vbChecked) Then
Check2.Value = False
Check3.Value = False
End If
End Sub

Private Sub Check2_Click()
If (Check2.Value = vbChecked) Then
Check1.Value = False
Check3.Value = False
End If
End Sub

Private Sub Check3_Click()
If (Check3.Value = vbChecked) Then
Check2.Value = False
Check1.Value = False
End If
End Sub

Private Sub Combo1_Click()
If (Combo1.ListIndex = 0) Then ''listýndex yerine "MARMARA BÖLGESÝ DE YAZILABÝLÝRDÝ"
Combo2.Clear
Combo2.AddItem ("Selam")
Combo2.AddItem ("Mehmet")
Combo2.AddItem ("Fatih")
End If

If (Combo1.ListIndex = 1) Then ''listýndex yerine "AKDENÝZ BÖLGESÝ DE YAZILABÝLÝRDÝ"
Combo2.Clear
Combo2.AddItem ("Banu")
Combo2.AddItem ("Ajdar")
Combo2.AddItem ("Adnan Þenses")
End If
End Sub


Private Sub Combo2_Click()
If (Combo2.Text = "Selam") Then
Combo3.AddItem ("Donaným")
Combo3.AddItem ("Yazýlým")
End If

If (Combo2.Text = "Mehmet") Then
Combo3.AddItem ("Grafik")
Combo3.AddItem ("Html")
End If

If (Combo2.Text = "Fatih") Then
Combo3.AddItem ("CSS")
Combo3.AddItem ("Java")
End If

If (Combo2.Text = "Banu") Then
Combo3.AddItem ("Müzik")
Combo3.AddItem ("Beden")
End If

If (Combo2.Text = "Ajdar") Then
Combo3.AddItem ("Matematik")
Combo3.AddItem ("Coðrafya")
End If

If (Combo2.Text = "Adnan Þenses") Then
Combo3.AddItem ("Harita")
Combo3.AddItem ("Ýklim")
End If

End Sub

Private Sub Command1_Click()
Text1.Text = turet()
End Sub


Private Sub Command2_Click()
If (Text2.Text = "") Then
MsgBox ("lütfen adý kýsmýný doldurunuz .")
Exit Sub
End If
If (Text3.Text = "") Then
MsgBox ("lütfen soyadý kýsmýný doldurunuz .")
Exit Sub
End If
If (Text4.Text = "") And (Text4.Text < 0) Or (Text4.Text > 100) Then
MsgBox ("Vize kýsmýný doldurunuz veya 0 ile 100 arasý sayý giriniz.")
Exit Sub
End If
If (Text5.Text = "") And (Text5.Text < 0) Or (Text5.Text > 100) Then
MsgBox ("Final kýsmýný doldurunuz veya 0 ile 100 arasý sayý giriniz.")
Exit Sub
End If
If (Text6.Text = "") And (Text6.Text < 0) Or (Text6.Text > 100) Then
MsgBox ("lütfen ödev kýsmýný doldurunuz veya 0 ile 100 arasý sayý giriniz.")
Exit Sub
End If

 ID = Text1.Text
 ad = Text2.Text
 soyad = Text3.Text
 vize = Text4.Text
 evize = vize
 vize = vize / 100 * 30
 final = Text5.Text
 efinal = final
 final = final / 100 * 50
 ödev = Text6.Text
 eödev = ödev
 ödev = ödev / 100 * 20
 bölüm = Combo1.Text
 hoca = Combo2.Text
 ders = Combo3.Text
 If (Check1.Value = vbChecked) Then
sinif = "Tekrar"
End If
If (Check2.Value = vbChecked) Then
sinif = "1.Sýnýf"
End If
If (Check3.Value = vbChecked) Then
sinif = "2.Sýnýf"
End If


ortalama = vize + final + ödev
If (ortalama >= 50) Then
durum = "SINIFI GEÇTÝ"
End If
If (ortalama < 50) Then
durum = "SINIFTA KALDI"
End If
If (ortalama >= 0) And (ortalama <= 35) Then
harfnotu = "FF"
End If

If (ortalama >= 36) And (ortalama <= 49) Then
harfnotu = "DC"
End If
If (ortalama >= 50) And (ortalama <= 64) Then
harfnotu = "CC"
End If
If (ortalama >= 65) And (ortalama <= 74) Then
harfnotu = "BC"
End If
If (ortalama >= 75) And (ortalama <= 90) Then
harfnotu = "BA"
End If
If (ortalama >= 91) And (ortalama <= 10) Then
harfnotu = "AA"
End If


cik = MsgBox("Hesaplama Gerçekleþsin mi ? ", 4 + 32 + 256, "Hesapla")
If (cik = 6) Then
List1.AddItem ("ID =" & ID & " " & "AD =" & ad & " " & "SOYAD =" & soyad & " " & "Vize =" & evize)
''List1.AddItem ("AD =" & ad)
''List1.AddItem ("SOYAD =" & soyad)
''List1.AddItem ("Vize =" & evize)
List1.AddItem ("Final =" & efinal & " " & "Ödev =" & eödev & " " & "Ortalama" & ortalama & " " & "Durum" & durum & " " & "Harf Not =" & harfnotu & " ")
''List1.AddItem ("Ödev =" & eödev)
''List1.AddItem ("Ortalama" & ortalama)
''List1.AddItem ("Durum" & durum)
''List1.AddItem ("Harf Not =" & harfnotu)
List1.AddItem ("Bölüm =" & bölüm & " " & "Öðretim Elemaný =" & hoca & " " & "Ders =" & ders & " " & "Sýnýf =" & sinif)
''List1.AddItem ("Öðretim Elemaný =" & hoca)
''List1.AddItem ("Ders =" & ders)
''List1.AddItem ("Sýnýf =" & sinif)
soru = MsgBox("Yedek.txt Dosyasýna Kayýt Yapýlsýn mý ? ", 4 + 32 + 256, "Kayýt")
If (soru = 6) Then
''kayýt yapma kodlarý


Open "yedek.txt" For Append As #1 '' append doysya kayýt yapar.


Print #1, ID, ad, soyad, ders, hoca, bölüm, durum, sinif, harfnotu, evize, efinal, eödev, ortalama '' yazdýrmak dosyaya
Close #1
MsgBox ("dosyaya baþarýlý bir þekilde kayýt yapýlmýþtýr.")
temizle
End If ''soru eveti endi
If (soru = 7) Then
temizle
End If ''soru hayýr edi

End If '' if sonu
If (cik = 7) Then

End If ''hayýr sonu


End Sub

Private Sub Form_Load()
Combo1.AddItem ("Bilgisayar Bölümü")
Combo1.AddItem ("Fotoðrafçýlýk Bölümü")
Text1.Text = turet()
Form1.acik = True


dosya = Dir("yedek.txt") ''val olup olmamasýný kontrol eden komut "Dir"
If Len(dosya) = 0 Then ''bu dosya yoksa
MsgBox ("Proje klasörüde yedek.txt dosyasý yoktur .")
soru = MsgBox("yedek.txt dosyasý yaratýlsýnmý", 4 + 32, "dosya yaratma iþlemi")
If (soru = 6) Then
Open "yedek.txt" For Output As #1 '' birinci dosya oluþturmak için #1
Close #1
MsgBox ("yedek.txt adýnda dosya oluþturuldu")
End If
Else
MsgBox ("Proje klasörüde yedek.txt dosyasý vardýr .")
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.acik = False
End Sub



Private Sub Text2_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
Text3.SetFocus
Exit Sub
End If
End Sub



Private Sub Text3_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
Text2.SetFocus
Exit Sub
End If
End Sub



Private Sub Text4_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 8) Then
If (KeyAscii < Asc("0")) Or (KeyAscii > Asc("9")) Then
KeyAscii = 0
End If
End If
End Sub



Private Sub Text5_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 8) Then
If (KeyAscii < Asc("0")) Or (KeyAscii > Asc("9")) Then
KeyAscii = 0
End If
End If
End Sub



Private Sub Text6_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 8) Then
If (KeyAscii < Asc("0")) Or (KeyAscii > Asc("9")) Then
KeyAscii = 0
End If
End If
End Sub

