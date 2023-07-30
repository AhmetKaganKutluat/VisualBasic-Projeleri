VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   8640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "ÇIKIÞ"
      Height          =   975
      Left            =   5880
      Picture         =   "Form1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2160
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SEÇ VE EKLE"
      Height          =   975
      Left            =   3360
      TabIndex        =   5
      Top             =   2160
      Width           =   2055
   End
   Begin VB.CheckBox Check3 
      Caption         =   "SADECE ÝLK 10 SAYI"
      Height          =   255
      Left            =   3360
      TabIndex        =   4
      Top             =   1560
      Width           =   2295
   End
   Begin VB.CheckBox Check2 
      Caption         =   "ÇÝFT SAYILAR ADET SAYILARI"
      Height          =   255
      Left            =   3360
      TabIndex        =   3
      Top             =   1080
      Width           =   2655
   End
   Begin VB.CheckBox Check1 
      Caption         =   "ÜÇE  BÖLÜNENLER"
      Height          =   255
      Left            =   3360
      TabIndex        =   2
      Top             =   600
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3360
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.ListBox List1 
      Height          =   6300
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sayilar, toplam As Integer


Private Sub Check1_Click()
If (Check1.Value = vbChecked) Then
Check2.Value = False
Check3.Value = False

End If
End Sub

Private Sub Check2_Click()
If (Check2.Value = vbChecked) Then
Check1.Value = False
Check3.Value = False

End If
End Sub

Private Sub Check3_Click()
If (Check3.Value = vbChecked) Then
Check1.Value = False
Check2.Value = False

End If
End Sub

Private Sub Command1_Click()
Combo1.Clear
If (Check1.Value = vbChecked) Then
For i = 0 To 30 Step 1
If (i Mod 3 = 0) Then
Combo1.AddItem (i)
End If
Next i
Exit Sub
End If  ''1.olay

If (Check2.Value = vbChecked) Then
For i = 1 To 30 Step 1
If (i Mod 2 = 0) Then
toplam = toplam + 1
End If
Next i
Combo1.AddItem (toplam)
Exit Sub
End If

If (Check3.Value = vbChecked) Then
For i = 0 To 30 Step 1
Combo1.AddItem (i)
If (i = 10) Then
MsgBox ("10 dan sonra eklenemez")
End
End If
Next i
Exit Sub
End If

End Sub

Private Sub Command2_Click()
Dim secim As Integer
secim = MsgBox("ÇIKIÞ YAPMAK ÝSTÝYORUMUSUZ", vbYesNo, "ÇIKIÞ")
If (secim = 6) Then
End
Else
Form2.Show 1
End If

End Sub

Private Sub Form_Load()
For i = 0 To 30
List1.AddItem (i)
Next i
End Sub
