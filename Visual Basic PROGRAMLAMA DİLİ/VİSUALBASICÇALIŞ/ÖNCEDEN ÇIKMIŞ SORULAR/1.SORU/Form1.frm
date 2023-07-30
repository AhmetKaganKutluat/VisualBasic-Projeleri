VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   5940
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
Dim sayilar(5) As Integer
Dim buyuk, i As Integer
For i = 0 To 4
sayilar(i) = InputBox(i + 1 & " . sayýyý giriniz")
If (sayilar(i) > buyuk) Then
buyuk = sayilar(i)
End If
Next i

MsgBox ("EN BÜYÜK SAYI " & buyuk)
End Sub

