VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form5 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "TEXT DOSYASINA KAYIT VE BU DOSYADAN OKUMA"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13560
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   162
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   13560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "SÝSTEMDEKÝ BÝLGÝLERÝ YÜKLE"
      Height          =   735
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6720
      Width           =   8295
   End
   Begin VB.ListBox List1 
      Height          =   2700
      Left            =   4920
      TabIndex        =   8
      Top             =   3840
      Width           =   3975
   End
   Begin RichTextLib.RichTextBox RichTextBox3 
      Height          =   2655
      Left            =   9000
      TabIndex        =   7
      Top             =   3840
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   4683
      _Version        =   393217
      BackColor       =   16777152
      Enabled         =   0   'False
      TextRTF         =   $"Form5.frx":0000
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "SÝSTEMDEKÝ BÝLGÝLERÝ YÜKLE"
      Height          =   1455
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2040
      Width           =   8295
   End
   Begin RichTextLib.RichTextBox RichTextBox2 
      Height          =   1695
      Left            =   4920
      TabIndex        =   5
      Top             =   240
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   2990
      _Version        =   393217
      BackColor       =   12640511
      Enabled         =   0   'False
      TextRTF         =   $"Form5.frx":0086
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   1080
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      Caption         =   "BÝLGÝLER KAYDET"
      Height          =   7215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0FF&
         Caption         =   "BÝLGÝLER.TXT DOSYASINA  KAYDET"
         Height          =   1215
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   5160
         Width           =   3375
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   3375
         Left            =   360
         TabIndex        =   3
         Top             =   1560
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   5953
         _Version        =   393217
         TextRTF         =   $"Form5.frx":010C
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   3375
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Dim soru As String '' bilgiler.txt ile ilgili dosyalar
Dim dosya As String '' bilgiler.txt ile ilgili dosyalar
Sub sil()
Text1.Text = ""
Text2.Text = ""
RichTextBox1.Text = ""
End Sub


Private Sub Command1_Click()
If (Text1.Text = "") Then
MsgBox ("MÜÞTERÝNÝN ADI GÝRÝLMEDÝ.")


Exit Sub
End If
If (Text2.Text = "") Then
MsgBox ("MÜÞTERÝNÝN SOYADI GÝRÝLMEDÝ.")
Exit Sub
End If
If (RichTextBox1.Text = "") Then
MsgBox ("MÜÞTERÝNÝN ADRESÝNÝ GÝRÝLMEDÝ.")
Exit Sub
End If

soru = MsgBox(UCase(Trim(Text1)) + " " + UCase(Trim(Text2)) + "ÝSÝMLÝ BÝLGÝ KAYIT YAPILSINMI ", 4 + 32, " DÝKKAT")
If (soru = 6) Then
Open "bilgiler.txt" For Append As #1 '' append doysya kayýt yapar.
ADI = Trim(UCase(Text1.Text))
SOYADI = Trim(UCase(Text2.Text))
ADRES = Trim(UCase(RichTextBox1.Text))
Print #1, ADI, SOYADI, ADRES '' yazdýrmak dosyaya
Close #1
MsgBox ("dosyaya baþarýlý bir þekilde kayýt yapýlmýþtýr.")
sil
Else
MsgBox ("KAYIT ÝÞEMÝNDEN VAZGEÇÝLDÝ")
sil
End If
End Sub

Private Sub Command2_Click()
yukle = MsgBox("SÝSTEM KAYITLI BÝLGÝLERÝ TEXT DOSYASINDAN ALINSINMI ?", 4 + 48, "DOSYA YÜKLEME")
If (yukle = 6) Then
RichTextBox2.LoadFile ("bilgiler.txt") ''dosyayý çekmek için
End If
End Sub

Private Sub Command3_Click()
Open "bilgiler.txt" For Input As #1 ''dosya okuma modunda okunan bilgileri deðiþkene aktaracaðý için deðiþken girdi olarak alýr.
Do Until EOF(1) ''dosyanýn sonuna kadar yap
Input #1, okunan ''bilgiler alýnýp okunan deðiþkenine atýlýr.
List1.AddItem (okunan)
RichTextBox3 = RichTextBox3.Text & vbTab & okunan
Loop '' do until bitiþi
Close #1



End Sub

Private Sub Form_Load()
dosya = Dir("bilgiler.txt") ''val olup olmamasýný kontrol eden komut "Dir"
If Len(dosya) = 0 Then ''bu dosya yoksa
MsgBox ("Proje klasörüde bilgiler.txt dosyasý yoktur .")
soru = MsgBox("bilgiler.txt dosyasý yaratýlsýnmý", 4 + 32, "dosya yaratma iþlemi")
If (soru = 6) Then
Open "bilgiler.txt" For Output As #1 '' birinci dosya oluþturmak için #1
Close #1
MsgBox ("bilgiler.txt adýnda dosya oluþturuldu")
End If
Else
MsgBox ("Proje klasörüde bilgiler.txt dosyasý vardýr .")
End If
End Sub

