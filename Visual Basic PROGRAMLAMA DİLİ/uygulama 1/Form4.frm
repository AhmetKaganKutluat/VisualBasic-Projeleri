VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "ALAN DENEME"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8910
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
   ScaleHeight     =   5490
   ScaleWidth      =   8910
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   4320
      TabIndex        =   7
      Text            =   "Text3"
      Top             =   4440
      Width           =   2895
   End
   Begin VB.ListBox List1 
      Height          =   1740
      Left            =   720
      TabIndex        =   6
      Top             =   2400
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "dikdötgen alan  hesapla"
      Height          =   1095
      Left            =   4920
      TabIndex        =   5
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "kare alan hesapla"
      Height          =   1095
      Left            =   4920
      TabIndex        =   4
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "kýsa kenar"
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "uzun kenar"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ukenar, kkenar, alan, cevre As Integer
Private Sub Command1_Click()
ukenar = Text1.Text
kkenar = Text2.Text
alan = 4 * ukenar
cevre = 4 * kkenar
List1.AddItem ("kare alan " & alan)
List1.AddItem ("kare cevre " & cevre)



End Sub

Private Sub Command2_Click()
ukenar = Text1.Text
kkenar = Text2.Text
alan = ukenar * kkenar
cevre = ((2 * ukenar) + (2 * kkenar))
List1.AddItem ("dikdortgen alan " & alan)
List1.AddItem ("dikdörtgen cevre " & cevre)

End Sub
