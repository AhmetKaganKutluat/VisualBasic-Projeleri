VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form8"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9045
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   162
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   9045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "ISTE"
      Height          =   1335
      Left            =   6600
      TabIndex        =   0
      Top             =   2640
      Width           =   2295
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form8.Cls
''ad = InputBox("L�tfen Ad�n�z� Yaz�n�z", "�sim giri�i")
Form8.CurrentX = 750
Form8.CurrentY = 500
Form8.FontBold = True
Form8.FontSize = 16
Form8.FontItalic = True
Form8.FontName = "Arial Tur"
''Print ad
secim = MsgBox("SE��M�N�Z� BEL�RT�N�Z", 35, "SE�")
If (secim = 6) Then
Print "evet t�kland�"
End If

If (secim = 7) Then
Print "hay�r t�kland�"
End If

If (secim = 2) Then
Print "iptal t�kland�"
End If
End Sub

