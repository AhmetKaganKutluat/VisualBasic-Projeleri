VERSION 5.00
Begin VB.Form Form7 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12420
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   162
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   12420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "DOMAÝN TÝPÝ"
      Height          =   4215
      Left            =   5760
      TabIndex        =   8
      Top             =   120
      Width           =   6375
      Begin VB.OptionButton Option5 
         Caption         =   "SERBEST ORGANÝZASYON"
         Height          =   375
         Left            =   480
         TabIndex        =   13
         Top             =   3480
         Width           =   5175
      End
      Begin VB.OptionButton Option4 
         Caption         =   "HÜKÜMET ORGANI"
         Height          =   375
         Left            =   480
         TabIndex        =   12
         Top             =   2760
         Width           =   5055
      End
      Begin VB.OptionButton Option3 
         Caption         =   "ASKERÝ KURUM"
         Height          =   375
         Left            =   480
         TabIndex        =   11
         Top             =   2040
         Width           =   5175
      End
      Begin VB.OptionButton Option2 
         Caption         =   "TÝCARÝ ÝÞLETME"
         Height          =   375
         Left            =   480
         TabIndex        =   10
         Top             =   1440
         Width           =   5415
      End
      Begin VB.OptionButton Option1 
         Caption         =   "EÐÝTÝM KURUMLARI"
         Height          =   375
         Left            =   480
         TabIndex        =   9
         Top             =   840
         Width           =   5415
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "HESAP OLUÞTUR"
      Height          =   1095
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2040
      Width           =   3135
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   3480
      Width           =   5175
   End
   Begin VB.TextBox Text3 
      Height          =   360
      Left            =   2040
      TabIndex        =   2
      Top             =   1440
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      Height          =   360
      Left            =   2040
      TabIndex        =   1
      Top             =   840
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Left            =   2040
      TabIndex        =   0
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label Label3 
      Caption         =   "DOMAÝN ADI :"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "KUL.SOYADI :"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "KUL.ADI :"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sonek As String
Private Sub Command1_Click()
If Option1.Value = True Then sonek = ".edu.tr"
If Option2.Value = True Then sonek = ".com.tr"
If Option3.Value = True Then sonek = ".mil.tr"
If Option4.Value = True Then sonek = ".gov.tr"
If Option5.Value = True Then sonek = ".org.tr"
Text4 = LCase(Trim(Text1)) & LCase(Trim(Text2)) & "@" & LCase(Trim(Text3)) & sonek  ''lcase çünkü  domain hesaparý küçük harf olurdu.
End Sub


