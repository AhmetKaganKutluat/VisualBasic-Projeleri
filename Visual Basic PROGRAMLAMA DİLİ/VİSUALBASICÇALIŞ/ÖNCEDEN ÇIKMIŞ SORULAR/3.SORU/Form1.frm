VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2280
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3945
   LinkTopic       =   "Form1"
   ScaleHeight     =   2280
   ScaleWidth      =   3945
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "SAYILARI �STEY�N�Z"
      Height          =   1575
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim kacsay�, sayilar(99), i As Integer
Dim toplam As Double
Dim ortalama As Double

kacsay� = InputBox("KA� SAYI G�R�L�CEK")

For i = 1 To kacsay�
sayilar(i) = InputBox(i & ". SAYIYI G�R�N�Z ")
toplam = toplam + sayilar(i)
Next i
ortalama = toplam / kacsay�

MsgBox ("SAYILARIN TOPLAMI " & toplam & " SAYILARIN ORTALALAMASI " & ortalama)
End Sub
