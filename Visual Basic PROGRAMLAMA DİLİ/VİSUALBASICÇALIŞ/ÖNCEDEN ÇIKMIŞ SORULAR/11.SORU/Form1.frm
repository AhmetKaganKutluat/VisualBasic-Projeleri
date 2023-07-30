VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4200
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "GÖNDER"
      Height          =   1215
      Left            =   840
      TabIndex        =   3
      Top             =   2640
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   1920
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   1200
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Command1_Click()
bil.adý = Text1.Text
bil.soyadý = Text2.Text
bil.adres = Text3.Text
Form2.Text1.Text = (bil.adý & bil.soyadý & bil.adres)
Form2.Show 1


End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
Text2.SetFocus
Exit Sub
End If


If (KeyAscii <> 8) Then
If (KeyAscii < Asc("a")) Or (KeyAscii > Asc("z")) Then
If (KeyAscii < Asc("A")) Or (KeyAscii > Asc("Z")) Then
KeyAscii = 0
End If
End If
End If
End Sub




Private Sub Text2_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
Text3.SetFocus
Exit Sub
End If


If (KeyAscii <> 8) Then
If (KeyAscii < Asc("a")) Or (KeyAscii > Asc("z")) Then
If (KeyAscii < Asc("A")) Or (KeyAscii > Asc("Z")) Then
KeyAscii = 0
End If
End If
End If

End Sub



Private Sub Text3_KeyPress(KeyAscii As Integer)

If (KeyAscii = 13) Then
Command1.SetFocus
Exit Sub
End If

If (KeyAscii <> 8) Then
If (KeyAscii < Asc("a")) Or (KeyAscii > Asc("z")) Then
If (KeyAscii < Asc("A")) Or (KeyAscii > Asc("Z")) Then
KeyAscii = 0
End If
End If
End If

End Sub
