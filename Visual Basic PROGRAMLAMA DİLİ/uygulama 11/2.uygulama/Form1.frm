VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3990
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   9165
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "KAÇ TANE ELEMAN ALINSIN ?"
      Height          =   1095
      Left            =   5400
      TabIndex        =   0
      Top             =   480
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dizi() As Integer ' boyutsuz bir dizi tanýmladýk '


Private Sub Command1_Click()
kactane = Val(InputBox("Kac Tane Sayi Almak Istersiniz ?"))
ReDim dizi(kactane) As Integer
'daha önceden tanýmlanan dizi yeniden yapýlandýrýldý'
For i = 1 To kactane
dizi(i) = Val(InputBox(Str(i) + " . Sayiyi Girin"))
toplam = toplam + dizi(i)

Next i
ortalama = toplam / kactane


Cls
For i = 1 To kactane
Print "Girilen " + Str(i) + " . sayi = " + Str(dizi(i))
Next
Print "Toplam   = " + Str(toplam)
Print "Ortalama = " + Str(ortalama)



End Sub
