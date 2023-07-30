VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "FUNCTÝON OLAYLARI"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10410
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
   ScaleHeight     =   4860
   ScaleWidth      =   10410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   480
      TabIndex        =   3
      Top             =   1200
      Width           =   3135
   End
   Begin VB.Frame Frame2 
      Caption         =   "KUVVET HESAPLA"
      Height          =   4935
      Left            =   6360
      TabIndex        =   1
      Top             =   0
      Width           =   4095
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0FF&
         Caption         =   "KUVVET HESAPLA"
         Height          =   2775
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   600
         Width           =   3615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "ALAN VE ÇEVRE"
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      Begin VB.CommandButton Command1 
         Caption         =   "HESAPLA"
         Height          =   975
         Left            =   480
         TabIndex        =   5
         Top             =   3120
         Width           =   3135
      End
      Begin VB.TextBox Text1 
         Height          =   615
         Left            =   480
         TabIndex        =   2
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00808080&
         Height          =   615
         Left            =   480
         TabIndex        =   4
         Top             =   2280
         Width           =   3135
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function hesapla(a, b)
alan = a * b
cevre = 2 * (a + b)
hesapla = ("ALAN:" & alan & "ÇEVRE :" & cevre)
End Function

Private Sub Command1_Click()
a = Val(Text1.Text)
b = Val(Text2.Text)
Label1.Caption = hesapla(a, b)
Text1.Text = ""
Text2.Text = ""
End Sub
Function kuvvet(sayi, us)
kuvvet = 1
For i = 1 To us
kuvvet = kuvvet * sayi
Next i
End Function
Private Sub Command2_Click()
On Error Resume Next ''hata varsa sýradakine geç
sayi = InputBox("TABAN SAYÝSÝNÝ GÝRÝNÝZ  :", "SAYI GÝRÝÞÝ")
us = InputBox("ÜST KUVVETÝNÝ GÝRÝNÝZ  :", "SAYI GÝRÝÞÝ")
MsgBox ("SAYININ KUVVETÝ  : " & kuvvet(sayi, us))


End Sub
