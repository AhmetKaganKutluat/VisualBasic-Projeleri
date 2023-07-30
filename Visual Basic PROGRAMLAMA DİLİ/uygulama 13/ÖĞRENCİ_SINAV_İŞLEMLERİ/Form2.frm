VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   1950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7920
   LinkTopic       =   "Form2"
   ScaleHeight     =   1950
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   1508
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   150
      Left            =   6960
      Top             =   1320
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   0
      Top             =   1320
      Width           =   3615
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 2
Label1.Caption = "PROGRAM YÜKLENÝYOR BEKLEYÝNÝZ  % " & ProgressBar1.Value
If (ProgressBar1.Value = 100) Then
Timer1.Enabled = False
Unload Me
MDIForm1.Show

End If
End Sub
