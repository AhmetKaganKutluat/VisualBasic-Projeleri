VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5025
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9555
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   9555
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Text Kutusu Sayi Giriþi"
      Height          =   735
      Left            =   2880
      TabIndex        =   0
      Top             =   2640
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sayi As Integer
Dim fakt As Integer
Dim toplam As Integer

Function hesapla(sayi)
          
          fakt = 1
          For i = 1 To sayi
          fakt = fakt * i
          Next i
          



End Function

Sub kendinekadar(sayi) ''procedure

For i = 1 To sayi
         toplam = toplam + i
          Next i


End Sub

Private Sub Command1_Click()
sayi = Text1.Text
If (sayi > 10) Or (sayi < 1) Then
MsgBox ("Sayi 1 ile 10 arasýnda olmalýdýr.")
Text1.Text = " "
Exit Sub
End If
hesapla (sayi)
kendinekadar (sayi)

MsgBox ("Sayinin faktöriyeli =" & fakt & Chr(13) & "Sayilarin toplami = " & toplam)




End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If (KeyAscii < 48 Or KeyAscii > 58) Then
KeyAscii = 0
End If
End If
End Sub
