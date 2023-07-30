VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4350
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   ScaleHeight     =   4350
   ScaleWidth      =   9465
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   960
      TabIndex        =   1
      Top             =   600
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "5 elemanlý dizi sayýlarýný al"
      Height          =   1095
      Left            =   6480
      TabIndex        =   0
      Top             =   840
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sayi(5) As Integer

Private Sub Command1_Click()
For i = 1 To 5
sayi(i) = Val(InputBox(Str(i) + " . sayýyý girin"))


Next
For i = 1 To 5
For j = 1 To 5

If (sayi(j) > sayi(i)) Then
gecici = sayi(i)

''5       8      7     3      6

sayi(i) = sayi(j)
sayi(j) = gecici

End If
Next
Next
List1.Clear


For i = 1 To 5
List1.AddItem sayi(i)
Next



End Sub
