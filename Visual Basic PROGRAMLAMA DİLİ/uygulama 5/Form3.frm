VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00FF8080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "TYPE MODUL ÝÞLEMÝ 2"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8505
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   162
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   8505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "TYPE MODUL 2 ÝÞLEMÝ"
      Height          =   615
      Left            =   3240
      TabIndex        =   3
      Top             =   240
      Width           =   3615
   End
   Begin VB.TextBox Text3 
      Height          =   360
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   6495
   End
   Begin VB.TextBox Text2 
      Height          =   360
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
bilgi.ADI = Trim(UCase(Text1.Text))
bilgi.SOYADI = Trim(UCase(Text2.Text))
bilgi.ADRES = Trim(UCase(Text3.Text))

Form4.Text1.Text = (bilgi.ADI + " " + bilgi.SOYADI + " " + bilgi.ADRES)
Form4.Show
End Sub
