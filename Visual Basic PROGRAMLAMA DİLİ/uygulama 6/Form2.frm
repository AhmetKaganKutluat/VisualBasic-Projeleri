VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   1470
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10470
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   10470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   480
      Top             =   960
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   1508
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   960
      Width           =   6615
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 2

For i = 0 To ProgressBar1.Value
Label1.Caption = "program yükleniyor   % " & i
Next i

If (ProgressBar1.Value >= 100) Then
Form1.Show

Timer1.Enabled = False
ProgressBar1.Value = 0
End If
End Sub
