VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0FF&
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11655
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   162
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5205
   ScaleWidth      =   11655
   Begin VB.ComboBox Combo2 
      Height          =   420
      Left            =   1800
      TabIndex        =   1
      Text            =   "Combo2"
      Top             =   1200
      Width           =   8175
   End
   Begin VB.ComboBox Combo1 
      Height          =   420
      ItemData        =   "Form1.frx":0000
      Left            =   1800
      List            =   "Form1.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   480
      Width           =   8175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Combo1_Click()
If (Combo1.ListIndex = 0) Then ''listýndex yerine "MARMARA BÖLGESÝ DE YAZILABÝLÝRDÝ"
Combo2.Clear
Combo2.AddItem ("ÝSTANBUL")
Combo2.AddItem ("EDÝRNE")
Combo2.AddItem ("KOCAELÝ")
Combo2.AddItem ("BURSA")
End If

If (Combo1.ListIndex = 1) Then ''listýndex yerine "AKDENÝZ BÖLGESÝ DE YAZILABÝLÝRDÝ"
Combo2.Clear
Combo2.AddItem ("MERSÝN")
Combo2.AddItem ("ADANA")
Combo2.AddItem ("ANTALYA")
Combo2.AddItem ("BURDUR")
End If

If (Combo1.ListIndex = 2) Then ''listýndex yerine "EGE BÖLGESÝ DE YAZILABÝLÝRDÝ"
Combo2.Clear
Combo2.AddItem ("ÝZMÝR")
Combo2.AddItem ("MANÝSA")
Combo2.AddItem ("AYDIN")
Combo2.AddItem ("MUÐLA")
End If




End Sub



Private Sub Form_Load()
Combo1.AddItem ("MARMARA BÖLGESÝ")
Combo1.AddItem ("AKDENÝZ BÖLGESÝ")
Combo1.AddItem ("EGE BÖLGESÝ")

End Sub

Private Sub Form_Unload(Cancel As Integer) ''çýkýþ yaparken
MDIForm1.Toolbar1.Buttons(1).Enabled = True
End Sub
