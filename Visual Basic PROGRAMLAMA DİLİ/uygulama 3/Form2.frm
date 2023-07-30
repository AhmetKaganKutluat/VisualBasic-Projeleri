VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11265
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   162
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   11265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   10440
      Top             =   2520
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Height          =   1215
      Left            =   2040
      TabIndex        =   0
      Top             =   600
      Width           =   7215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sayilar(5) As Integer
Private Sub Form_Load()
For i = 1 To 5
sayilar(i) = Val(InputBox(i & " . Sayiyi giriniz ", "SAYI GÝRÝÞÝ")) '' VAL YERÝNE CÝNT GÝREBÝLÝRZ ÇÜMKÜ INPUTBOX STRÝNG BÝR DEÐERDÝR
topla = topla + sayilar(i)
Next i
ortalama = topla / 5
Label1.Caption = ("TOPLAM : " & topla & " ORTALAMA  " & ortalama)
End Sub

Private Sub Timer1_Timer()
Form2.Caption = ("DÝZÝLÝ INPUTBOX ÝÞLEM ÖRNEKLEK UYGULAMASI " & Now)
End Sub
