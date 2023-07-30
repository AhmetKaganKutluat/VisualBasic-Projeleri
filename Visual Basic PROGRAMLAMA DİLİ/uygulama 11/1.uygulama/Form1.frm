VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3510
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   ScaleHeight     =   3510
   ScaleWidth      =   7950
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim isim(2) As String
Dim notlar(3) As Byte
Private Sub Form_Click()
isim(0) = "Ahmet Kagan"
isim(1) = "Furkan Zeki"
isim(2) = "Adam Otu"
notlar(0) = 25
notlar(1) = 45
notlar(2) = 35
notlar(3) = 15
Print isim(0), notlar(0)
Print isim(1), notlar(1)
Print isim(2), notlar(2)
Print "Görsel Programlama"; notlar(3)



End Sub


