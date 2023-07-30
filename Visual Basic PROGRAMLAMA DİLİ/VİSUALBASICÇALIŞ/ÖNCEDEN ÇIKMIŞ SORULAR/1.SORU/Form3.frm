VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   2040
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4110
   LinkTopic       =   "Form3"
   ScaleHeight     =   2040
   ScaleWidth      =   4110
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "SORGULU ÇIKIÞ"
      Height          =   975
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   2775
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   240
      Top             =   0
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
secim = MsgBox("PROGRAM KAPANSINMI ", 4 + 32, "ÇIKIÞ")
If (secim = 6) Then
End
Exit Sub
End If
If (secim = 7) Then
Me.Hide
Form1.Show
Exit Sub
End If
End Sub

Private Sub Timer1_Timer()
Form3.Caption = Now
End Sub
