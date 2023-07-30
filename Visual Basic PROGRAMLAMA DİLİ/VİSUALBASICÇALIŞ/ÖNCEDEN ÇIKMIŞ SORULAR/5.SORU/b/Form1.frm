VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4065
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6810
   LinkTopic       =   "Form1"
   ScaleHeight     =   4065
   ScaleWidth      =   6810
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Faktöriyel için týkla"
      Height          =   855
      Left            =   1800
      TabIndex        =   0
      Top             =   960
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sayi As Integer
Public fakt As Integer
Function faktoriyel(sayi)
For i = 1 To sayi
fakt = fakt * i
Next i
Form2.Caption = "Faktoriyel=" & fakt
Form2.Label1.Caption = "Sayinin faktoriyeli = " & fakt
Form2.Show
Unload Me

End Function

Private Sub Command1_Click()
fakt = 1
sayi = Val(InputBox("Sayiyi girin = ", "Faktöriyel"))
faktoriyel (sayi)

End Sub



          


