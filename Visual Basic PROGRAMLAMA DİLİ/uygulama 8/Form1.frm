VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "DOSYAYA KAYIT VE DOSYADAN OKUMA"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12345
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   162
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   12345
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "DOSYAYA KAYDET"
      Height          =   1095
      Left            =   6600
      TabIndex        =   3
      Top             =   360
      Width           =   3495
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   735
      Left            =   600
      TabIndex        =   2
      Top             =   1680
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   1296
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form1.frx":0000
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   960
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub temizle()
Text1.Text = ""
Text2.Text = ""
RichTextBox1.Text = ""

End Sub

Private Sub Command1_Click()
If (Text1.Text = "") Then
MsgBox ("MÜÞTERÝNÝN ADINI YAZMADINIZ LÜTFEN YAZINIZ ")
Exit Sub
End If
If (Text2.Text = "") Then
MsgBox ("MÜÞTERÝNÝN SOYADINI YAZMADINIZ LÜTFEN YAZINIZ ")
Exit Sub
End If
If (RichTextBox1.Text = "") Then
MsgBox ("MÜÞTERÝNÝN ADRESÝNÝ YAZMADINIZ LÜTFEN YAZINIZ ")
Exit Sub
End If

soru = MsgBox(UCase(Trim(Text1.Text)) + " " + UCase(Trim(Text2.Text)) + " isimi kayýt yapýlsýnmý ?  ", 4 + 32 + 256, "sisteme kayýt ")

If (soru = 6) Then '' dosya oluþturup kayýt yapma
Open "kayit.txt" For Append As #1
adi = Trim(UCase(Text1.Text))
soyadi = Trim(UCase(Text2.Text))
adres = Trim(UCase(RichTextBox1.Text))
Print #1, adi, soyadi, adres
Close #1
MsgBox ("bilgiler sisteme kayýt yapýlmýþtýr")
temizle
Else
temizle ''procedure

End If '' if sonu

End Sub
