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
      Caption         =   "S�STEMDEK� B�LG�LER� Y�KLE"
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
      Caption         =   "S�STEMDEK� B�LG�LER� Y�KLE"
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
      Caption         =   "B�LG�LER KAYDET"
      Height          =   7215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0FF&
         Caption         =   "B�LG�LER.TXT DOSYASINA  KAYDET"
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
MsgBox ("M��TER�N�N ADI G�R�LMED�.")


Exit Sub
End If
If (Text2.Text = "") Then
MsgBox ("M��TER�N�N SOYADI G�R�LMED�.")
Exit Sub
End If
If (RichTextBox1.Text = "") Then
MsgBox ("M��TER�N�N ADRES�N� G�R�LMED�.")
Exit Sub
End If

soru = MsgBox(UCase(Trim(Text1)) + " " + UCase(Trim(Text2)) + "�S�ML� B�LG� KAYIT YAPILSINMI ", 4 + 32, " D�KKAT")
If (soru = 6) Then
Open "bilgiler.txt" For Append As #1 '' append doysya kay�t yapar.
ADI = Trim(UCase(Text1.Text))
SOYADI = Trim(UCase(Text2.Text))
ADRES = Trim(UCase(RichTextBox1.Text))
Print #1, ADI, SOYADI, ADRES '' yazd�rmak dosyaya
Close #1
MsgBox ("dosyaya ba�ar�l� bir �ekilde kay�t yap�lm��t�r.")
sil
Else
MsgBox ("KAYIT ��EM�NDEN VAZGE��LD�")
sil
End If
End Sub

Private Sub Command2_Click()
yukle = MsgBox("S�STEM KAYITLI B�LG�LER� TEXT DOSYASINDAN ALINSINMI ?", 4 + 48, "DOSYA Y�KLEME")
If (yukle = 6) Then
RichTextBox2.LoadFile ("bilgiler.txt") ''dosyay� �ekmek i�in
End If
End Sub

Private Sub Command3_Click()
Open "bilgiler.txt" For Input As #1 ''dosya okuma modunda okunan bilgileri de�i�kene aktaraca�� i�in de�i�ken girdi olarak al�r.
Do Until EOF(1) ''dosyan�n sonuna kadar yap
Input #1, okunan ''bilgiler al�n�p okunan de�i�kenine at�l�r.
List1.AddItem (okunan)
RichTextBox3 = RichTextBox3.Text & vbTab & okunan
Loop '' do until biti�i
Close #1



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

