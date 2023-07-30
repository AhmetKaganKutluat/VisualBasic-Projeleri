VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1365
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4170
   LinkTopic       =   "Form1"
   ScaleHeight     =   1365
   ScaleWidth      =   4170
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "GÖSTER"
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim sayi As String
sayi = Text1.Text
If (sayi < 1) Or (sayi > 7) Then
MsgBox ("YANLIÞ DEÐER GÝRDÝNÝZ")
Text1.Text = ""
Else
Select Case sayi
Case 1:
MsgBox ("pazartesi")
Text1.Text = ""
Case 2:
MsgBox ("salý")
Text1.Text = ""
Case 3:
MsgBox ("çarþamba")
Text1.Text = ""
Case 4:
MsgBox ("perþembe")
Text1.Text = ""
Case 5:
MsgBox ("cuma")
Text1.Text = ""
Case 6:
MsgBox ("cumartesi")
Text1.Text = ""
Case 7:
MsgBox ("pazar")
Text1.Text = ""
End Select
End If
End Sub
