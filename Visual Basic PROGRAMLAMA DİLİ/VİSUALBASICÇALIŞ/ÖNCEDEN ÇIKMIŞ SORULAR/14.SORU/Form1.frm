VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   4410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   ScaleHeight     =   4410
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1800
      Top             =   2640
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   4080
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   840
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
Cls
If (Combo1.Text = "0") Then
Print ("0 a bastýn bastýn")
End If



If (Combo1.Text = "2") Then

Print (1 * 2)
Print (2 * 2)
Print (3 * 2)
Print (4 * 2)
Print (5 * 2)
Print (6 * 2)
Print (7 * 2)
Print (8 * 2)
Print (9 * 2)
Print (10 * 2)
End If

If (Combo1.Text = "4") Then

Print (1 * 4)
Print (2 * 4)
Print (3 * 4)
Print (4 * 4)
Print (5 * 4)
Print (6 * 4)
Print (7 * 4)
Print (8 * 4)
Print (9 * 4)
Print (10 * 4)
End If


If (Combo1.Text = "6") Then

Print (1 * 6)
Print (2 * 6)
Print (3 * 6)
Print (4 * 6)
Print (5 * 6)
Print (6 * 6)
Print (7 * 6)
Print (8 * 6)
Print (9 * 6)
Print (10 * 6)
End If


If (Combo1.Text = "8") Then

Print (1 * 8)
Print (2 * 8)
Print (3 * 8)
Print (4 * 8)
Print (5 * 8)
Print (6 * 8)
Print (7 * 8)
Print (8 * 8)
Print (9 * 8)
Print (10 * 8)
End If


If (Combo1.Text = "10") Then

Print (1 * 10)
Print (2 * 10)
Print (3 * 10)
Print (4 * 10)
Print (5 * 10)
Print (6 * 10)
Print (7 * 10)
Print (8 * 10)
Print (9 * 10)
Print (10 * 10)
End If

End Sub

Private Sub Form_Load()
For i = 0 To 10 Step 2 '' 2 þer 2 þer arttýrdýk çiftleri ekledik.
Combo1.AddItem (i)
Next i


End Sub


