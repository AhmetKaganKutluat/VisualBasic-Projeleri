VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H008080FF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8385
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   8385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   2520
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "KONTROL ET"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   1
      Top             =   240
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   720
      TabIndex        =   2
      Top             =   1080
      Width           =   5175
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim isim As String
Dim sayi As Integer

Private Sub Command1_Click()
sayi = Text1.Text
isim = Text2.Text
If (sayi < 50) Then
Label1.Caption = isim & " " & "SINIFTA KALDI"
Else
Label1.Caption = isim & " " & "SINIFI GEÇTÝ"
End If
End Sub




Private Sub Text1_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 8) Then
If (KeyAscii < Asc("0")) Or (KeyAscii > Asc("9")) Then
KeyAscii = 0
End If
End If
End Sub



Private Sub Text2_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 8) Then
If (KeyAscii < Asc("a")) Or (KeyAscii > Asc("z")) Then
If (KeyAscii < Asc("A")) Or (KeyAscii > Asc("Z")) Then
KeyAscii = 0
End If
End If
End If
End Sub
