VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3735
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8385
   LinkTopic       =   "Form1"
   ScaleHeight     =   3735
   ScaleWidth      =   8385
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List2 
      Height          =   2010
      Left            =   5520
      TabIndex        =   6
      Top             =   360
      Width           =   2055
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   3240
      TabIndex        =   5
      Top             =   360
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SÝL"
      Height          =   735
      Left            =   960
      TabIndex        =   4
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "EKLE"
      Height          =   735
      Left            =   960
      TabIndex        =   3
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   1320
      Width           =   2055
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   480
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   840
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   480
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label1 
      Height          =   495
      Left            =   3240
      TabIndex        =   7
      Top             =   2640
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
 Combo2.Clear
        If (Combo1.Text = "AKDENÝZ") Then
            Combo2.AddItem ("ANTALYA")
            Combo2.AddItem ("MERSÝN")
            Combo2.AddItem ("ÝSTANBUL")
            Exit Sub
            End If
        If (Combo1.Text = "MARMARA") Then
            Combo2.AddItem ("ÝSTANBUL")
            Combo2.AddItem ("BURSA")
            Combo2.AddItem ("EDÝRNE")
            Exit Sub
            End If
        If (Combo1.Text = "EGE") Then
            Combo2.AddItem ("ÝZMÝR")
            Combo2.AddItem ("MANÝSA")
            Combo2.AddItem ("AYDIN")
            Exit Sub
        End If
End Sub
Private Sub Command1_Click()
Dim plaka, sehir, veri As String
Dim soru As Integer

If (Text1.Text = "") Then
MsgBox ("boþ býrakýlamaz")
Exit Sub
End If

veri = Trim(Text1.Text)
plaka = Strings.Left(veri, 2)
sehir = Strings.Right(veri, Strings.Len(veri) - 2)
soru = MsgBox(veri & "KAYIT YAPILSINMI", 4 + 64, "KAYIT ÝÞLEM")
If (soru = 6) Then
List1.AddItem (plaka)
List2.AddItem (sehir)
Text1.Text = ""
Label1.Caption = "TOPLAM KAYIT:" + CStr(List1.ListCount)
Exit Sub
End If
If (soru = 7) Then
Text1.Text = ""
Combo1.Clear
Combo2.Clear
Label1.Caption = ""
Text1.SetFocus




Exit Sub
End If
End Sub

Private Sub Form_Load()
Dim i As Integer
For i = 1 To 3
Combo1.AddItem (Choose(i, "AKDENÝZ", "MARMARA", "EGE"))
Next i

End Sub

