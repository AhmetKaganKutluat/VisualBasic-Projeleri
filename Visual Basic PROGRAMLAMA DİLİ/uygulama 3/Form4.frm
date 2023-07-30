VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "FIBONACCÝ PROCEDURLE"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5400
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   162
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "KAÇ TANE FÝBO SAYISI GÝRÝLSÝN"
      Height          =   1215
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5040
      Width           =   3975
   End
   Begin VB.ListBox List1 
      Height          =   4380
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fibo(500) As Integer
Sub hesap(sayi) ''procedure
sayi = 10
fibo(1) = 1
For i = 2 To sayi
fibo(i) = fibo(i - 1) + fibo(i - 2)
Next i
For i = 1 To sayi
List1.AddItem (i & ".  fibo sayýsý : " & fibo(i))
Next i
End Sub

Private Sub Command1_Click()
List1.Clear
sayi = Val(InputBox("KAÇ TANE FÝBO SAYISI GÖSTERÝLSÝN ", "FÝBO SERÝSÝ"))
hesap (sayi)
End Sub


