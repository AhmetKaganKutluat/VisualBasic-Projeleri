VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Form2"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8475
   LinkTopic       =   "Form2"
   ScaleHeight     =   4725
   ScaleWidth      =   8475
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option4 
      Caption         =   "BÖLME"
      Height          =   255
      Left            =   3360
      TabIndex        =   8
      Top             =   2760
      Width           =   1335
   End
   Begin VB.OptionButton Option3 
      Caption         =   "ÇARPMA"
      Height          =   255
      Left            =   3360
      TabIndex        =   7
      Top             =   2040
      Width           =   1335
   End
   Begin VB.OptionButton Option2 
      Caption         =   "ÇIKAR"
      Height          =   255
      Left            =   3360
      TabIndex        =   6
      Top             =   1320
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "TOPLA"
      Height          =   255
      Left            =   3360
      TabIndex        =   5
      Top             =   720
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "LÜTFEN SEÇÝNÝZ ..."
      Height          =   3015
      Left            =   3240
      TabIndex        =   4
      Top             =   360
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "HESAPLA"
      Height          =   975
      Left            =   360
      TabIndex        =   3
      Top             =   1800
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFC0&
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   2880
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim s1, s2, sonuc As Integer
s1 = CInt(Text1.Text)
s2 = CInt(Text2.Text)


If (Option1.Value = True) Then
sonuc = s1 + s2
Text3.Text = ("SONUC = " & sonuc)
Exit Sub
End If

If (Option2.Value = True) Then
sonuc = s1 - s2
Text3.Text = ("SONUC = " & sonuc)
Exit Sub
End If

If (Option3.Value = True) Then
sonuc = s1 * s2
Text3.Text = ("SONUC = " & sonuc)
Exit Sub
End If

If (Option4.Value = True) Then
sonuc = s1 / s2
Text3.Text = ("SONUC = " & sonuc)
Exit Sub
End If



End Sub
