VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4545
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8025
   LinkTopic       =   "Form1"
   ScaleHeight     =   4545
   ScaleWidth      =   8025
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   5175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "2 x 10  dizi elemanlarýnýn karesi"
      Height          =   1455
      Left            =   5400
      TabIndex        =   0
      Top             =   3000
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dizi(2, 5) As Integer

Private Sub Command1_Click()
On Error Resume Next 'hata varsa sýradankýne gec'
For i = 1 To 5
dizi(1, i) = InputBox(Str(i) + " . elemaný girin ")
dizi(2, i) = dizi(1, i) * dizi(1, i)
Next

For i = 1 To 5
''Print Str(dizi(1, i)) + " sayýnýn karesi = " + Str(dizi(2, i))
List1.AddItem Str(dizi(1, i)) + " Sayýnýn karesi = " + Str(dizi(2, i))
Next




End Sub
