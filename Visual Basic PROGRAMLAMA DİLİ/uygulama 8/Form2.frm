VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   4005
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7395
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   162
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   4005
   ScaleWidth      =   7395
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "ÝSÝMLERÝ EKLE"
      Height          =   1215
      Left            =   840
      TabIndex        =   2
      Top             =   1800
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ÝSÝMLERÝ ÝSTE"
      Height          =   1095
      Left            =   840
      TabIndex        =   1
      Top             =   480
      Width           =   3015
   End
   Begin VB.ListBox List1 
      Height          =   3180
      Left            =   4200
      TabIndex        =   0
      Top             =   360
      Width           =   2775
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim isimler(6) As String
Private Sub Command1_Click()
For i = 1 To 6
isimler(i) = InputBox(i & " . ismi giriniz ")
Next i


End Sub

Private Sub Command2_Click()
For i = 1 To 6
List1.AddItem (isimler(i))
Next i
End Sub
