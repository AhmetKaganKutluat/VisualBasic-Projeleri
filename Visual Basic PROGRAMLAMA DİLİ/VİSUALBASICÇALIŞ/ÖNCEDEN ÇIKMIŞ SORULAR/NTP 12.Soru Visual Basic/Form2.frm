VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form2"
   ClientHeight    =   8715
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   20760
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   20760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5760
      Top             =   3840
   End
   Begin VB.Frame Frame2 
      Caption         =   "Soru 2"
      Height          =   2415
      Left            =   6480
      TabIndex        =   16
      Top             =   2160
      Width           =   13935
      Begin VB.CheckBox Check8 
         Caption         =   "�ok Kat�l�yorum."
         Height          =   375
         Left            =   6480
         TabIndex        =   26
         Top             =   960
         Width           =   1575
      End
      Begin VB.CheckBox Check7 
         Caption         =   "Kat�l�yorum"
         Height          =   495
         Left            =   4560
         TabIndex        =   25
         Top             =   840
         Width           =   1215
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Kat�lm�yorum."
         Height          =   255
         Left            =   2640
         TabIndex        =   24
         Top             =   960
         Width           =   1335
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Hi� Kat�lm�yorum."
         Height          =   255
         Left            =   360
         TabIndex        =   23
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "�retici Firman�n Sat�� Politikas�ndan Memnunum."
         Height          =   375
         Left            =   360
         TabIndex        =   22
         Top             =   240
         Width           =   5775
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Soru 1"
      Height          =   1935
      Left            =   6480
      TabIndex        =   15
      Top             =   120
      Width           =   13935
      Begin VB.CheckBox Check4 
         Caption         =   "�ok kat�l�yorum"
         Height          =   375
         Left            =   6600
         TabIndex        =   21
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Kat�l�yorum"
         Height          =   375
         Left            =   4560
         TabIndex        =   20
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Kat�lm�yorum"
         Height          =   495
         Left            =   3000
         TabIndex        =   19
         Top             =   960
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Hi� Kat�lm�yorum"
         Height          =   375
         Left            =   600
         TabIndex        =   18
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Yap�lan Sat�� Sonras� Servisten Memnunum."
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Sisteme Y�kle"
      Height          =   1215
      Left            =   360
      TabIndex        =   14
      Top             =   3360
      Width           =   5055
   End
   Begin VB.ListBox List1 
      Height          =   3180
      Left            =   120
      TabIndex        =   13
      Top             =   4680
      Width           =   20295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Yinele"
      Height          =   735
      Left            =   2760
      TabIndex        =   12
      Top             =   120
      Width           =   2295
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   1320
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   2760
      Width           =   1815
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1320
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1320
      MaxLength       =   25
      TabIndex        =   9
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1320
      MaxLength       =   25
      TabIndex        =   8
      Top             =   1320
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1320
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Sisteme Y�kl� Kay�t Say�s� = "
      Height          =   255
      Left            =   480
      TabIndex        =   27
      Top             =   8160
      Width           =   3015
   End
   Begin VB.Label Label6 
      Caption         =   "Yaiad��� �l�e = "
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Ya�ad��� il ="
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "M��teri Soyad� ="
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "M��teri Ad� ="
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Anket�r ="
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Anket No ="
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
   Begin VB.Menu mnformat 
      Caption         =   "i�lemler"
      Begin VB.Menu mnsil 
         Caption         =   "Sil"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim durum As Integer
Dim anketno As Variant
Dim anket�r, soruanket, m��teriad, m��terisoyad, il, il�e As String
Sub temizle()

''List1.Clear
For Each Control In Me.Controls
If TypeName(Control) = "TextBox" Then
Control.Text = ""
End If
Next

For Each Control In Me.Controls
If TypeName(Control) = "ComboBox" Then
Control.ListIndex = -1 'or control.clear
End If
Next


Check1.Value = False
Check2.Value = False
Check3.Value = False
Check4.Value = False
Check5.Value = False
Check6.Value = False
Check7.Value = False
Check8.Value = False
Text1.Text = turet()

End Sub

Function turet()
Dim sifre As String ''text kutusuna aktar�lan �ifremdir
Dim karakterler(25) ''ka� tane karakter varsa o kdr ekle
karakterler(0) = "A"
karakterler(1) = "B"
karakterler(2) = "T"
karakterler(3) = "6"
karakterler(4) = "8"
karakterler(5) = "C"
karakterler(6) = "S"
karakterler(7) = "F"
karakterler(8) = "7"
karakterler(9) = "D"
karakterler(10) = "P"
karakterler(11) = "I"
karakterler(12) = "M"
karakterler(13) = "Y"
karakterler(14) = "1"
karakterler(15) = "H"
karakterler(16) = "2"
karakterler(17) = "X"
karakterler(18) = "3"
karakterler(19) = "4"
karakterler(20) = "9"
karakterler(21) = "W"
karakterler(22) = "T"
karakterler(23) = "7"
karakterler(24) = "J"
karakterler(25) = "H"
For i = 1 To 6 ''  basamak say�s�
Randomize
sifre = sifre & karakterler(Int(Rnd() * 25))
Next i
turet = sifre
End Function
''18.47





Private Sub Check1_Click()
If (Check1.Value = vbChecked) Then  ''check 1 se�iliyse bunlar� yap dedik
Check2.Value = False
Check3.Value = False
Check4.Value = False
Check5.Value = False
Check6.Value = False
Check7.Value = False
Check8.Value = False
End If
End Sub

Private Sub Check2_Click()
If (Check2.Value = vbChecked) Then  ''check 1 se�iliyse bunlar� yap dedik
Check1.Value = False
Check3.Value = False
Check4.Value = False
Check5.Value = False
Check6.Value = False
Check7.Value = False
Check8.Value = False
End If
End Sub

Private Sub Check3_Click()
If (Check3.Value = vbChecked) Then  ''check 1 se�iliyse bunlar� yap dedik
Check2.Value = False
Check1.Value = False
Check4.Value = False
Check5.Value = False
Check6.Value = False
Check7.Value = False
Check8.Value = False
End If
End Sub

Private Sub Check4_Click()
If (Check4.Value = vbChecked) Then  ''check 1 se�iliyse bunlar� yap dedik
Check2.Value = False
Check3.Value = False
Check1.Value = False
Check5.Value = False
Check6.Value = False
Check7.Value = False
Check8.Value = False
End If
End Sub

Private Sub Check5_Click()
If (Check5.Value = vbChecked) Then  ''check 1 se�iliyse bunlar� yap dedik
Check2.Value = False
Check3.Value = False
Check4.Value = False
Check1.Value = False
Check6.Value = False
Check7.Value = False
Check8.Value = False
End If
End Sub

Private Sub Check6_Click()
If (Check6.Value = vbChecked) Then  ''check 1 se�iliyse bunlar� yap dedik
Check2.Value = False
Check3.Value = False
Check4.Value = False
Check5.Value = False
Check1.Value = False
Check7.Value = False
Check8.Value = False
End If
End Sub

Private Sub Check7_Click()
If (Check7.Value = vbChecked) Then  ''check 1 se�iliyse bunlar� yap dedik
Check2.Value = False
Check3.Value = False
Check4.Value = False
Check5.Value = False
Check6.Value = False
Check1.Value = False
Check8.Value = False
End If
End Sub

Private Sub Check8_Click()
If (Check8.Value = vbChecked) Then  ''check 1 se�iliyse bunlar� yap dedik
Check2.Value = False
Check3.Value = False
Check4.Value = False
Check5.Value = False
Check6.Value = False
Check7.Value = False
Check1.Value = False
End If
End Sub

Private Sub Combo2_Click()
Combo3.Clear
If (Combo2.Text = "Ankara") Then
Combo3.Clear
Combo3.AddItem ("Sincan")
Combo3.AddItem ("Ke�i�ren")
Combo3.AddItem ("�ankaya")
End If

If (Combo2.Text = "�stanbul") Then
Combo3.Clear
Combo3.AddItem ("Kartal")
Combo3.AddItem ("Beyo�lu")
Combo3.AddItem ("Kad�k�y")
End If

If (Combo2.Text = "Antalya") Then
Combo3.Clear
Combo3.AddItem ("Serik")
Combo3.AddItem ("Manavgat")
Combo3.AddItem ("Titreyeng�l")
End If



End Sub

Private Sub Command1_Click()
Text1.Text = turet()
End Sub

Private Sub Command2_Click()

If (Check1.Value = vbChecked) Then  ''check 1 se�iliyse bunlar� yap dedik
durum = 1
End If
If (Check2.Value = vbChecked) Then  ''check 1 se�iliyse bunlar� yap dedik
durum = 2
End If
If (Check3.Value = vbChecked) Then  ''check 1 se�iliyse bunlar� yap dedik
durum = 3
End If
If (Check4.Value = vbChecked) Then  ''check 1 se�iliyse bunlar� yap dedik
durum = 4
End If
If (Check5.Value = vbChecked) Then  ''check 1 se�iliyse bunlar� yap dedik
durum = 1
End If
If (Check6.Value = vbChecked) Then  ''check 1 se�iliyse bunlar� yap dedik
durum = 2
End If
If (Check7.Value = vbChecked) Then  ''check 1 se�iliyse bunlar� yap dedik
durum = 3
End If
If (Check8.Value = vbChecked) Then  ''check 1 se�iliyse bunlar� yap dedik
durum = 4
End If
If (Check8.Value = vbChecked) Or (Check7.Value = vbChecked) Or (Check6.Value = vbChecked) Or (Check5.Value = vbChecked) Then

soruanket = "2. Soru"
End If
If (Check1.Value = vbChecked) Or (Check2.Value = vbChecked) Or (Check3.Value = vbChecked) Or (Check4.Value = vbChecked) Then

soruanket = "1. Soru"
End If

anketno = Text1.Text
anket�r = Combo1.Text
m��teriad = Text2.Text
m��terisoyad = Text3.Text
il = Combo2.Text
il�e = Combo3.Text

soru = MsgBox("Bilgiler Kaydedilsin mi ?", 4 + 32, "KAYIT")
If (soru = 6) Then
List1.AddItem ("Anket NO = " & anketno & " " & "Anket�r = " & anket�r & " " & "M��teri Ad = " & m��teriad & " " & "M��teri Soyad� = " & m��terisoyad & " " & "�li = " & il & " " & "�l�esi = " & il�e & " " & "Cevaplad��� Soru = " & soruanket & " " & "Kat�lma Durumu = " & durum)
temizle

End If '' evetin endi
      
If (soru = 7) Then

temizle

End If '' hay�r�n endi


End Sub ''butonun endi

Private Sub Form_Load()
Form1.ac�k = True
Text1.Text = turet()
Combo1.AddItem ("Mete Han")
Combo1.AddItem ("Cengiz Han")
Combo1.AddItem ("Timur Han")
Combo1.AddItem ("K�r�at Han")
Combo1.AddItem ("Alparslan Han")

Combo2.AddItem ("Ankara")
Combo2.AddItem ("�stanbul")
Combo2.AddItem ("Antalya")

End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.ac�k = False
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then ''farenizdeki sa� tu� t�klanm�� ise vbRightButton yerine "2"de yaz�labilir
PopupMenu mnformat  '' mnformat ismini verdi�imiz yer a��l�r not:i�i dolu olmas� laz�m ��nk� i� taraf gozukuyor iz burda i�ine g�ster ekledik.
End If
End Sub

Private Sub mnsil_Click()
intX = List1.ListIndex   '' se�ili indexi de�i�kene att�k
List1.RemoveItem (intX)  '' se�ili indexi sildik

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then  ''enter tu�una bas�ld���nda (13 ascii de enter demek)
Text3.SetFocus
Exit Sub
End If
If (KeyAscii <> 8) Then
If (KeyAscii < Asc("a")) Or (KeyAscii > Asc("z")) Then '' sadece silme tu�u ve harfler �al���r onun do��nda giri� yap�l�rsa yok say demek
If (KeyAscii < Asc("A")) Or (KeyAscii > Asc("Z")) Then
KeyAscii = 0
End If
End If
End If
End Sub



Private Sub Text3_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then  ''enter tu�una bas�ld���nda (13 ascii de enter demek)
Text2.SetFocus ''text2 ye focusla imleci
Exit Sub
End If
If (KeyAscii <> 8) Then
If (KeyAscii < Asc("a")) Or (KeyAscii > Asc("z")) Then '' sadece silme tu�u ve harfler �al���r onun do��nda giri� yap�l�rsa yok say demek
If (KeyAscii < Asc("A")) Or (KeyAscii > Asc("Z")) Then
KeyAscii = 0
End If
End If
End If
End Sub

Private Sub Timer1_Timer()
Label9.Caption = ("Sisteme Y�kl� Kay�t Say�s� = " & List1.ListIndex + 1 & "/" & List1.ListCount)
End Sub


