VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ÞÝFRE TÜRETME"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10800
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   162
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   10800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ÞÝFRE TÜRET"
      Height          =   1455
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   4575
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFC0&
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   4575
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function turet()
Dim sifre As String ''text kutusuna aktarýlan þifremdir
Dim karakterler(25) ''kaç tane karakter varsa o kdr ekle
karakterler(0) = "A"
karakterler(1) = "B"
karakterler(2) = "T"
karakterler(3) = "6"
karakterler(4) = "8"
karakterler(5) = "{"
karakterler(6) = "S"
karakterler(7) = "F"
karakterler(8) = "7"
karakterler(9) = "+"
karakterler(10) = "P"
karakterler(11) = "?"
karakterler(12) = "M"
karakterler(13) = "m"
karakterler(14) = "r"
karakterler(15) = "H"
karakterler(16) = "w"
karakterler(17) = "X"
karakterler(18) = "/"
karakterler(19) = "*"
karakterler(20) = "#"
karakterler(21) = "&"
karakterler(22) = "t"
karakterler(23) = "7"
karakterler(24) = "j"
karakterler(25) = "H"
For i = 1 To 8 '' 8 basamalklý þifre
Randomize
sifre = sifre & karakterler(Int(Rnd() * 25))
Next i
turet = sifre
End Function

Private Sub Command1_Click()
Text1.Text = turet
End Sub
