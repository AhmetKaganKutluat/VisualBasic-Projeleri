VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   12465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   240
      Top             =   2640
   End
   Begin VB.CommandButton Command1 
      Caption         =   "HESAPLA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   960
      TabIndex        =   6
      Top             =   2280
      Width           =   2895
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   960
      MaxLength       =   2
      TabIndex        =   2
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   960
      MaxLength       =   2
      TabIndex        =   1
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   960
      MaxLength       =   2
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   9
      Top             =   2400
      Width           =   7335
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   8
      Top             =   4080
      Width           =   7335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   7
      Top             =   3240
      Width           =   7335
   End
   Begin VB.Label Label3 
      Caption         =   "C:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "B:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "A:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   480
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   1140
      Left            =   3240
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   600
      Width           =   5580
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, b, c, delta, tekkok, x1, x2 As Integer
Dim parabol As String
Private Sub Command1_Click()
a = Val(Text1.Text)
b = Val(Text2.Text)
c = Val(Text3.Text)

delta = (b * b) - (4 * a * c)
Label6.Caption = "DELTA SONUCU :" & delta

If (a < 0) Then
parabol = "AÞAÐI YÖNLÜ PARABOL "
Else
parabol = "YUKARI YÖNLÜ PARABOL"
End If

Label4.Caption = parabol

If (delta < 0) Then
Label5.Caption = "FONKSÝYONUN REEL KÖKÜ YOKTUR."
Else
If (delta = 0) Then
tekkok = (-1 * b) / (2 * a)
Label5.Caption = "TEK KÖK VARDIR . KÖK DEÐERÝ :" & tekkok
Else
If (delta > 0) Then
x1 = ((-1 * b) + Math.Sqr(delta)) / (2 * a) '' math.sgr karakök demek sgr karakök demektir.
x2 = ((-1 * b) - Math.Sqr(delta)) / (2 * a)
x1 = Math.Round(x1, 2) '' iki basamak yuvarladý
x2 = Math.Round(x2, 2)
Label5.Caption = ("ÇÝFT KÖK VARDIR X1 : " & x1 & " X2  :" & x2)
End If
End If
End If


End Sub



Private Sub Timer1_Timer()
If Trim(Text1.Text) = "" Or Trim(Text2.Text) = "" Or Trim(Text3.Text) = "" Then
Command1.Enabled = False
Else
Command1.Enabled = True
End If
End Sub
