VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Form3"
   ClientHeight    =   4650
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9015
   LinkTopic       =   "Form3"
   ScaleHeight     =   4650
   ScaleWidth      =   9015
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command7 
      Caption         =   "VERÝTABANI BÝLGÝLER.TXT DOSYASINI YEDEKLEME"
      Height          =   855
      Left            =   600
      TabIndex        =   6
      Top             =   2160
      Width           =   7095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "www.akdeniz.edu.tr"
      Height          =   855
      Left            =   5520
      TabIndex        =   5
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CommandButton Command5 
      Caption         =   "PAÝNT AÇ"
      Height          =   855
      Left            =   3000
      TabIndex        =   4
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "HESAP MAKÝNASI AÇ"
      Height          =   855
      Left            =   600
      TabIndex        =   3
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "PROJE DOSYA ÝÇÝNDE TEXT AÇ"
      Height          =   855
      Left            =   5520
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "EXCEL DOSYA AÇ"
      Height          =   855
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "WORD DOSYA AÇ"
      Height          =   855
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "www.akdeniz.edu.tr"
      Height          =   255
      Left            =   3480
      TabIndex        =   7
      Top             =   3840
      Width           =   1935
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim yedek As New FileSystemObject
Private Sub Command1_Click()
Shell "C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE", vbNormalFocus
End Sub

Private Sub Command2_Click()
Shell "C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE", vbNormalFocus
End Sub


Private Sub Command3_Click()
Shell "notepad bilgiler.txt", vbNormalFocus
End Sub

Private Sub Command4_Click()
Shell "calc.exe", vbNormalFocus
End Sub

Private Sub Command5_Click()
Shell "mspaint.exe", vbNormalFocus
End Sub

Private Sub Command6_Click()
Shell "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe www.akdeniz.edu.tr", vbNormalFocus

End Sub

Private Sub Command7_Click()
'' proje menüsünden referanslara geliyoruz . microsoft scripting runtime seçilir.
'' yukarý bi deðiþken atarýz genelde çalýþan
sor = MsgBox("BÝLGÝLER.TXT DOSYASININ YEDEÐÝNÝ ALMAK ÝSTÝYORMUSUNZ ? ", 4 + 32, "KOPYALAMA ÝÞLEMÝ ")
If sor = 6 Then
yedek.CopyFile "bilgiler.txt", "yedek.txt", True
End If
End Sub

Private Sub Label1_Click()
Shell "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe www.akdeniz.edu.tr", vbNormalFocus

End Sub
