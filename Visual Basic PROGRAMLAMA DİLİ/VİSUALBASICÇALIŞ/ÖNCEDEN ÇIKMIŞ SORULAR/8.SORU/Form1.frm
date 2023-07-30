VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3555
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   ScaleHeight     =   3555
   ScaleWidth      =   7440
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "EKLE"
      Height          =   1455
      Left            =   4800
      TabIndex        =   0
      Top             =   720
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim adý, soyadý As String
Dim sayilar(99), kactane, i, soru As Integer
Dim ortalama, toplam As Double

adý = InputBox("ADINIZI GÝRÝNÝZ")
soyadý = InputBox("SOYADI GÝRÝNÝZ")
kactane = InputBox("KAÇ TANE SAYI GÝRÝÞÝ OLUCAK")
For i = 1 To kactane
sayilar(i) = InputBox(i & ". SAYIYI GÝRÝNÝZ")
toplam = toplam + sayilar(i)
Next i
ortalama = toplam / kactane
soru = MsgBox("EKLEME YAPMAK ÝSTÝYORMUSUNUZ", vbYesNo, "SORU")
If (soru = 6) Then
List1.Visible = True
List1.AddItem (List1.ListCount + 1 & " ADI " & adý & " SOYADI " & soyadý & " TOPLAMI " & toplam & " ORTALAMA " & ortalama)
Else
MsgBox ("ÝYÝ GÜNLER")
End If
End Sub
