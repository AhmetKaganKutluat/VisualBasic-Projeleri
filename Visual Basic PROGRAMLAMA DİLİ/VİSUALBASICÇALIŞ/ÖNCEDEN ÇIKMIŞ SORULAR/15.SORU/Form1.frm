VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5745
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   ScaleHeight     =   5745
   ScaleWidth      =   7590
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   3000
      Top             =   3000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Öðrenci Not Hesapla = "
      Height          =   1095
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   2535
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   480
      Top             =   3000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public acik As Boolean
Private Sub Command1_Click()
Form2.Show 1
End Sub

Private Sub Timer1_Timer()
Form1.Caption = "Anasayfa    " & Now
End Sub

Private Sub Timer2_Timer()
If (acik = True) Then
Command1.Enabled = False
End If
If (acik = False) Then
Command1.Enabled = True
End If

End Sub
