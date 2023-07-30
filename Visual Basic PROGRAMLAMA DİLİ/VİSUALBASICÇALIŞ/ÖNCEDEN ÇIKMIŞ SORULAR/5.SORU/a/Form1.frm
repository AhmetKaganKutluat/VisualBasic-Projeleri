VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4395
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   ScaleHeight     =   4395
   ScaleWidth      =   6405
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "SAYI GÝRMEK ÝÇÝN BASIN"
      Height          =   855
      Left            =   1680
      TabIndex        =   0
      Top             =   1200
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sayi As Integer
Public fakt As Integer


Sub faktoriyel()
sayi = Val(InputBox("Sayiyi girin = ", "Faktöriyel"))
          
For i = 1 To sayi
fakt = fakt * i
Next i
Form2.Label1.Caption = "Sayinin faktoriyeli=" & fakt
Form2.Caption = "Faktoriyel = " & fakt
Form2.Show
Unload Me


End Sub
Private Sub Command1_Click()
fakt = 1
faktoriyel



          
End Sub
