VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   6750
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2055
      Left            =   1440
      TabIndex        =   1
      Top             =   1560
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   3625
      _Version        =   393217
      TextRTF         =   $"Form1.frx":0000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Y�KLEME"
      Height          =   1095
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
secim = MsgBox("VER� TABANINDAK� B�LG�LER S�STEME Y�KLENS�NM� ", 4 + 32, "Y�KLEME")
If (secim = 6) Then

RichTextBox1.LoadFile ("bilgiler.txt")

Exit Sub
End If
If (secim = 7) Then
MsgBox ("��LEM �PTAL ")
Exit Sub
End If
End Sub

Private Sub Form_Load()
dosya = Dir("bilgiler.txt") ''val olup olmamas�n� kontrol eden komut "Dir"
If Len(dosya) = 0 Then ''bu dosya yoksa
MsgBox ("Proje klas�r�de bilgiler.txt dosyas� yoktur .")
soru = MsgBox("bilgiler.txt dosyas� yarat�ls�nm�", 4 + 32, "dosya yaratma i�lemi")
If (soru = 6) Then
Open "bilgiler.txt" For Output As #1 '' birinci dosya olu�turmak i�in #1
Close #1
MsgBox ("bilgiler.txt ad�nda dosya olu�turuldu")
End If
Else
MsgBox ("Proje klas�r�de bilgiler.txt dosyas� vard�r .")
End If
End Sub

