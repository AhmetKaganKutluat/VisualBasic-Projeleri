VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form5 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "TEXT DOSYASINA KAYIT VE BU DOSYADAN OKUMA"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11775
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
   ScaleHeight     =   6945
   ScaleWidth      =   11775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   1080
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      Caption         =   "BÝLGÝLER KAYDET"
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.CommandButton Command1 
         Caption         =   "BÝLGÝLER.TXT DOSYASINA  KAYDET"
         Height          =   1215
         Left            =   360
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
         Enabled         =   -1  'True
         TextRTF         =   $"Form5.frx":0000
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

Private Sub Command1_Click()
Open "bilgiler.txt" For Append As #1 '' append doysya kayýt yapar.
ADI = Trim(UCase(Text1.Text))
SOYADI = Trim(UCase(Text2.Text))
ADRES = Trim(UCase(RichTextBox1.Text))
Print #1, ADI, SOYADI, ADRES '' yazdýrmak dosyaya
Close #1
MsgBox ("dosyaya baþarýlý bir þekilde kayýt yapýlmýþtýr.")
Text1.Text = ""
Text2.Text = ""
RichTextBox1.Text = ""
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
