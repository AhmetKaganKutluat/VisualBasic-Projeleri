VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "ÝSTE"
      Height          =   1335
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim sayilar(2), toplam, cýkarma, bolme, carpma As Integer
Dim secim As String
For i = 1 To 2
sayilar(i) = Val(InputBox(i & " . SAYÝYÝ GÝRÝNÝZ"))
Next i
secim = InputBox(" +    -  *  /  = BÝRÝNÝ SEÇÝNÝZ ")
Select Case secim

Case "+":
toplam = sayilar(1) + sayilar(2)
If (toplam < 0) Then
toplam = toplam * -1
End If
MsgBox ("SAYILARIN TOPLAMI " & toplam)

Case "-":
cýkarma = sayilar(1) - sayilar(2)
If (cýkarma < 0) Then
cýkarma = cýkarma * -1
End If
MsgBox ("SAYILARIN FARKI " & cýkarma)

Case "*":
carpma = sayilar(1) * sayilar(2)
If (carpma < 0) Then
carpma = carpma * -1
End If
MsgBox ("SAYILARIN ÇARPIMI " & carpma)
    

Case "/":
bolme = sayilar(1) / sayilar(2)
If (bolme < 0) Then
bolme = bolme * -1
End If
MsgBox ("SAYILARIN BÖLÜMÜ " & bolme)

End Select
End Sub
