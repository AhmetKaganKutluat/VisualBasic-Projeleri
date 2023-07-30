VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "MENÜ ÝÞLEMLER"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   11145
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   162
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   11145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "TOPLAMI VER"
      Height          =   975
      Left            =   8280
      TabIndex        =   5
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   5280
      TabIndex        =   4
      Top             =   3480
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   5280
      TabIndex        =   3
      Top             =   2640
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "MESAJI GÖSTER"
      Height          =   975
      Left            =   8160
      TabIndex        =   2
      Top             =   240
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   5160
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
   Begin VB.ListBox List1 
      Height          =   5100
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
   Begin VB.Menu mndosya 
      Caption         =   "DOSYA"
      Begin VB.Menu mnyeni 
         Caption         =   "YENÝ"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnac 
         Caption         =   "AÇ"
      End
      Begin VB.Menu mnkaydet 
         Caption         =   "Kaydet"
         Shortcut        =   ^S
      End
      Begin VB.Menu mntumunukaydet 
         Caption         =   "TÜMÜNÜ KAYDET"
      End
      Begin VB.Menu mncýzgý 
         Caption         =   "-"
      End
      Begin VB.Menu mncýk 
         Caption         =   "ÇIKIÞ"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mnduzen 
      Caption         =   "DÜZEN"
      Begin VB.Menu mnkes 
         Caption         =   "KES"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnkopyala 
         Caption         =   "KOPYALA"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnyapýþtýr 
         Caption         =   "YAPIÞTIR"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnçýzgý2 
         Caption         =   "-"
      End
      Begin VB.Menu mntumunusec 
         Caption         =   "TÜMÜNÜ SEÇ"
         Shortcut        =   ^B
      End
   End
   Begin VB.Menu mnformat 
      Caption         =   "FORMAT"
      Begin VB.Menu mnofýs 
         Caption         =   "OFÝS UYGULAMASI"
         Begin VB.Menu mneskýw 
            Caption         =   "WORD 2002 - 2003"
         End
         Begin VB.Menu mnyenýw 
            Caption         =   "WORD 2007 - 2016"
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
goster
End Sub

Private Sub Command2_Click()
toplama
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
''sað tuþla popap eklmek için burasý kullanýlr.
If Button = vbRightButton Then ''farenizdeki sað tuþ týklanmýþ ise vbRightButton yerine "2"de yazýlabilir
PopupMenu mnformat
End If
End Sub

Private Sub mncýk_Click()
End
''Unload Me ''buda çýkýþ için
End Sub
