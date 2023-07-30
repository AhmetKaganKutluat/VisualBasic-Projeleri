VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   ScaleHeight     =   3210
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   360
      Top             =   2040
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ÞÝFRE ONAYLA"
      Height          =   855
      Left            =   1440
      TabIndex        =   3
      Top             =   2040
      Width           =   2295
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "GÖSTER / GÝZLE"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   1440
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   1920
      Left            =   4320
      Picture         =   "Form1.frx":0000
      Top             =   360
      Width           =   1920
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ÞÝFRE                :"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "KULLANICI ADI :"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim kullanýcý, sifre As String
Dim hak As Integer

Private Sub Check1_Click()
If (Check1.Enabled = True) Then
Text2.PasswordChar = ""
Exit Sub
End If
If (Check1.Enabled = False) Then
Text2.PasswordChar = "*"
Exit Sub
End If
End Sub

Private Sub Command1_Click()


kullanýcý = Text1.Text
sifre = Text2.Text

If (kullanýcý = "furkan") And (sifre = "2020") Then
MDIForm1.StatusBar1.Panels.Add.Text = kullanýcý '' tam olmadý
Unload Me
Form2.Show

Else
hak = hak - 1
MsgBox ("þifre yada kullanýcý adý hatalý  hakkýnýz = " & hak)
If (hak = 0) Then
End  '' asd
End If
Text1.Text = ""
Text2.Text = ""
Exit Sub
End If
End Sub

Private Sub Form_Load()
hak = 3
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
Text2.SetFocus
End If
End Sub



Private Sub Text2_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
Command1.SetFocus
End If
End Sub
