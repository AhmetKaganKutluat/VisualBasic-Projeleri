VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   705
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
   Begin VB.Menu mnmenu 
      Caption         =   "MENU"
      Begin VB.Menu mngoster 
         Caption         =   "GÖSTER"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
PopupMenu mnmenu
End If
End Sub

Private Sub mngoster_Click()
MsgBox ("FURKAN GERDAN")
End Sub
