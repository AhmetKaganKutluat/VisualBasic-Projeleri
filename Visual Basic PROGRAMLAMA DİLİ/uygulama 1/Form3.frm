VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form3"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11850
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   162
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   11850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   540
      Left            =   720
      TabIndex        =   2
      Top             =   960
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   360
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "HESAPLA"
      Height          =   1095
      Left            =   4680
      TabIndex        =   0
      Top             =   360
      Width           =   3615
   End
   Begin VB.Label Label1 
      Height          =   735
      Left            =   720
      TabIndex        =   3
      Top             =   1920
      Width           =   6255
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s1, s2, sonuc As Integer

Private Sub Command1_Click()
''Label1.Caption = Val(Text1) + Val(Text2)
s1 = CInt(Text1.Text) ''sayýsal ifadeye çevirmek için val kullanýlrý.
s2 = CInt(Text2.Text)
sonuc = s1 + s2
Label1.Caption = "sayýlarýn toplamý :" & sonuc ''metin ve sayýlarda & iaþreti stringlerde + birþeþtirir.
End Sub
