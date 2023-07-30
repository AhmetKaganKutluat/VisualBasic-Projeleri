VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3945
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9150
   LinkTopic       =   "Form1"
   ScaleHeight     =   3945
   ScaleWidth      =   9150
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "6 elemanlý dizi sayýlarý al"
      Height          =   855
      Left            =   6360
      TabIndex        =   0
      Top             =   2640
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sayilar(6) As Integer

Private Sub Command1_Click()
For i = 1 To 6
sayilar(i) = Val(InputBox(Str(i) + " . sayiyi girin"))
If (sayilar(i) < 0) Then negatif_adet = negatif_adet + 1
If (sayilar(i) > 0) Then pozitif_adet = pozitif_adet + 1
If (sayilar(i) = 0) Then notr_adet = notr_adet + 1
Next
Cls

Print "Pozýtýf adet sayýsý = " + Str(pozitif_adet)
Print "Negatif adet sayýsý = " + Str(negatif_adet)
Print "notr adet sayýsý = " + Str(notr_adet)

End Sub
