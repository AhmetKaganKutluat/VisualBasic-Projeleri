VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "GÖRSEL PROGRAMLAMA 1"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   11490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   4320
      TabIndex        =   1
      Top             =   1320
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "GÖSTER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim metin As String
Private Sub Command1_Click()
metin = Trim(UCase(Text1.Text))
Form2.Show
Form2.Label1.Caption = metin

End Sub

    
