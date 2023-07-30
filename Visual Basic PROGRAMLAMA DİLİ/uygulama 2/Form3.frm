VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13575
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   13575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   7200
      Top             =   4560
   End
   Begin VB.ListBox List2 
      Height          =   4740
      Left            =   8640
      TabIndex        =   3
      Top             =   360
      Width           =   4575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "TÜMÜNÜ EKLE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   5400
      TabIndex        =   2
      Top             =   2400
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SEÇÝLENÝ EKLE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   5400
      TabIndex        =   1
      Top             =   600
      Width           =   2775
   End
   Begin VB.ListBox List1 
      Height          =   4740
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4575
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   6120
      Top             =   4560
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   5400
      Width           =   4575
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If (List1.ListIndex < 0) Then
MsgBox ("eklenecek ders adý seçilmedi lütfen seçiniz")
Exit Sub
End If

List2.AddItem List1.Text
List1.RemoveItem (List1.ListIndex)


End Sub

Private Sub Command2_Click()
For i = 0 To List1.ListCount
List2.AddItem List1.List(i)
Next i
List1.Clear
Label1.Caption = ""
Command2.Enabled = False
End Sub

Private Sub Form_Load()
List1.AddItem ("NESNE TABANLI PROGRAMLA")
List1.AddItem ("GÖRSEL  PROGRAMLA")
List1.AddItem ("MAT")
List1.AddItem ("TARÝH")
List1.AddItem ("BEDEN")
End Sub


Private Sub List1_Click()
Label1.Caption = ("DURUM" & List1.ListIndex + 1 & " / " & List1.ListCount)
End Sub



Private Sub Timer1_Timer()
Form3.Caption = ("GÖRSEL PROGRAMLAMA 1 " & Now)
End Sub

Private Sub Timer2_Timer()
If (List1.ListCount = 0) Then   '' count hesaplama listenin eleman sayýsý yoksa
Command1.Enabled = False
End If
End Sub





