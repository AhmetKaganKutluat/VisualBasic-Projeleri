VERSION 5.00
Begin VB.Form Form9 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form9"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8265
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   8265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim secim As Integer
Form9.Show
Form8.FontBold = True
Form8.FontSize = 16
Form8.FontItalic = True
Form8.FontName = "Arial Tur"

Print "1  - KAYIT G�R���"
Print "2  - KAYIT OKUMA"
Print "3  - KAYIT D�ZELTME"
Print "4  - KAYIT �IKI�"

secim = InputBox("LUTFEN 1- 4 ARASI SEC�M YAPINIZ", "SEC�M�N�Z")
If secim = 1 Then
GoTo giris
ElseIf secim = 2 Then
GoTo okuma
ElseIf secim = 3 Then
GoTo d�zeltme
ElseIf secim = 4 Then
GoTo cik
MsgBox ("1-aras� de�er giriniz")
Exit Sub
End If



giris:
MsgBox ("Kay�tlar�n giri�inin yap�ld��� yerdir ")
Exit Sub
okuma:
MsgBox ("kay�tlar� bilgiler okutuldugu yerdir.")
Exit Sub
d�zeltme:
MsgBox ("kay�tlar� bilgiler d�zeltildi�i yerdir.")
cik:
End
End Sub
