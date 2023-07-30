VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "TYPE ÖZELLÝÐÝ VE MODUL ÝÞLEMÝ"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9510
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   162
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "TYPE / MODUL ÇALIÞTIR"
      Height          =   1335
      Left            =   3960
      TabIndex        =   3
      Top             =   240
      Width           =   3375
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3495
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808080&
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   3240
      Width           =   7095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808080&
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2520
      Width           =   7095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   7095
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
ogr.ADI = Trim(UCase(Text1.Text))
ogr.SOYADI = Trim(UCase(Text2.Text))
ogr.ADRES = Trim(UCase(Text3.Text))

Label1.Caption = "ÖÐRENCÝNÝN ADI : " & ogr.ADI
Label2.Caption = "ÖÐRENCÝNÝN SOYADI : " & ogr.SOYADI
Label3.Caption = "ÖÐRENCÝNÝN ADRES : " & ogr.ADRES

End Sub
