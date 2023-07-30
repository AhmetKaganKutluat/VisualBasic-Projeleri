VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "GENEL ÝÞLEMLER"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11625
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   11625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000E&
      Caption         =   "SÝSTEMÝ KAPAT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8640
      TabIndex        =   0
      Top             =   3360
      Width           =   2655
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
''kapatmanýn 1. yönemi '' Unload Me ''aktif formu kapatýr
''End ''tüm programý,projeyi sonlandýrýr.
cik = MsgBox("program kapatýlsýn mý ", 4 + 32 + 256, "çýkýþ")
If (cik = 6) Then

End ''çýkma kodu
End If '' if sonu

End Sub

