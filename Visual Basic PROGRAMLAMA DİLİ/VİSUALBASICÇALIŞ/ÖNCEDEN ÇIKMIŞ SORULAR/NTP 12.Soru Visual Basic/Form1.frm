VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   12135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   3120
      Top             =   3240
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1440
      Top             =   3000
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   1335
      Left            =   3120
      TabIndex        =   1
      Top             =   240
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Anket Yükle"
      Height          =   1335
      Left            =   240
      Picture         =   "Form1.frx":0000
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public acýk As Boolean
Private Sub Command1_Click()
Form2.Show 1
End Sub

Private Sub Timer1_Timer()
Form1.Caption = "Anket Programý    " & Now
End Sub

Private Sub Timer2_Timer()
If (acýk = True) Then
Command1.Enabled = False
End If
If (acýk = False) Then
Command1.Enabled = True
End If

End Sub
