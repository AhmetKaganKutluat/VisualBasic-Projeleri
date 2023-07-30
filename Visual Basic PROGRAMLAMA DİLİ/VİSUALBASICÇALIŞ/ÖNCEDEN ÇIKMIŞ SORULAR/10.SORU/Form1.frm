VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5340
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   ScaleHeight     =   5340
   ScaleWidth      =   9510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Faktöriyel Hesapla"
      Height          =   855
      Left            =   3120
      TabIndex        =   0
      Top             =   1680
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sayi, i As Integer
Public fakt As Integer
Private Sub Command1_Click()
sayi = InputBox("Faktöriyeli buluncak sayý giriniz.", "Faktöriyel")
i = 1
fakt = 1

Do
i = i + 1
fakt = fakt * i
Loop While i < sayi
''MsgBox ("Faktöriyel = " & fakt & " i =" & i)
Form2.Text1.Text = "Sayýnýn Faktöriyeli = " & fakt
Form2.Show 1

End Sub
