VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   11145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "DO IT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   480
      Width           =   3855
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   600
      MaxLength       =   150
      TabIndex        =   4
      Top             =   2520
      Width           =   6975
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   2760
      MaxLength       =   5
      TabIndex        =   3
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   600
      MaxLength       =   5
      TabIndex        =   2
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   600
      MaxLength       =   20
      TabIndex        =   1
      Top             =   1080
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   600
      MaxLength       =   20
      TabIndex        =   0
      Top             =   360
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim secim As Variant    ''variant genel de�i�ken demek
Dim ad, soyad, adres As String
Dim gelir, gider, fark As Integer

Private Sub Command1_Click()
If (Text1.Text = "") Then
MsgBox ("l�tfen ad� k�sm�n� doldurunuz .")
Exit Sub
End If
If (Text2.Text = "") Then
MsgBox ("l�tfen soyad� k�sm�n� doldurunuz .")
Exit Sub
End If
If (Text3.Text = "") Then
MsgBox ("gelir k�sm�n� doldurunuz.")
Exit Sub
End If
If (Text4.Text = "") Then
MsgBox ("gider k�sm�n� doldurunuz.")
Exit Sub
End If
If (Text5.Text = "") Then
MsgBox ("l�tfen adres k�sm�n� doldurunuz .")
Exit Sub
End If
git:
secim = InputBox("L�TFEN A�A�IDAK� SE��MLERDEN B�R�N� YAPINIZ", "L�tfen Se�iniz", " 1 - 9 , HESAPLA ,5,S ")
secim = UCase(Trim(secim)) ''b�y�k harfe �evirmek i�in

Select Case secim
Case 1 To 9: MsgBox ("g�rsel programlama")
GoTo git
Case S: End
Case "5": MsgBox ("SELAM KACMAZ " & vbCr & "MANAVGAT MYO")
Case "HESAPLA":
ad = Trim(UCase(Text1))
soyad = Trim(UCase(Text2))
gelir = CInt(Text3)  ''cint ordaki de�eri say�sala d�k ifadesidir.
gider = CInt(Text4)
adres = Trim(UCase(Text5))
fark = gelir - gider
fark = fark * -1
MsgBox (ad & " " & soyad & fark & " tl maa�l� " & vbCr & adres & "adresinde ya�amaktad�r.")
GoTo git
Case Else
MsgBox ("tan�ms�z  tu�lama  ")
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
End Select
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)

If (KeyAscii = 13) Then
Text4.SetFocus
Exit Sub
End If

If KeyAscii <> 8 Then
If (KeyAscii < 48 Or KeyAscii > 58) Then
KeyAscii = 0
End If
End If
End Sub



Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If (KeyAscii < 48 Or KeyAscii > 58) Then
KeyAscii = 0
End If
End If
End Sub
