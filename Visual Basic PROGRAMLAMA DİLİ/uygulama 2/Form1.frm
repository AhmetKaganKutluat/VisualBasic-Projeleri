VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "M��TER� ��LEM PANEL�"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   11190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "Form1.frx":0000
      Left            =   480
      List            =   "Form1.frx":0002
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1080
      Width           =   2895
   End
   Begin VB.CommandButton Command3 
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
      Height          =   1215
      Left            =   3600
      TabIndex        =   3
      Top             =   240
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   480
      TabIndex        =   2
      Top             =   600
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   480
      TabIndex        =   1
      Top             =   240
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "GENEL ��LEMLER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command2_Click()
Form2.Show 1 ''1 koyarsak arkadaki forma ge�meyi engeller .

End Sub

Private Sub Command3_Click()
'' BU 1. Y�NTEM

''If (Trim(Text1.Text) = "") Then
''MsgBox ("M��TER�N�N ADINI YAZINIZ .")
''Exit Sub '' d�ng�y� burda bitirir.
''End If
''If (Trim(Text2.Text) = "") Then
''MsgBox ("M��TER�N�N SOYADINI YAZINIZ .")
''Exit Sub
''End Sub


''BU 2.Y�NTEM

''If (Trim(Text1.Text) = "") Or (Trim(Text2.Text) = "") Then
''MsgBox ("M��TER�N�N ADINI VEYA SOYADINI YAZMADINIZ ")
''End If
''End Sub

''BU 3.Y�NTEM
If (Trim(Text1.Text) = "") Then
MsgBox ("L�TFEN M��TER�N�N ADINI YAZINIZ" & Chr(13) & "                   HATA YAPTINIZ")
Else
If (Trim(Text2.Text) = "") Then
MsgBox ("L�TFEN M��TER�N�N SOYADINI YAZINIZ" & Chr(13) & "                   HATA YAPTINIZ")
Else
If (Combo1.Text = "") Then
MsgBox ("l�tfen bir isim se�iniz")
Else
MsgBox (" BA�ARILI ")
End If
End If
End If
End Sub

Private Sub Form_Load()
'' 1. Y�NTEM
''Combo1.AddItem ("selam ka�maz")
''Combo1.AddItem ("mustafa")
''Combo1.AddItem ("derya")
''Combo1.AddItem ("hasan")
''Combo1.AddItem ("furkan")

'' 2. Y�NTEM
For i = 1 To 5
Combo1.AddItem Choose(i, "nesne programlam", "g�rsel programlama", "mat", "tarih", "beden")
Next i ''for sonu
End Sub






