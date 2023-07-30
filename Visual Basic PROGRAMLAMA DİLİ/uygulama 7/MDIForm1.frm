VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   5250
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11625
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   1164
      ButtonWidth     =   2752
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "müþteri iþlemleri"
            Object.ToolTipText     =   "müþteri bilgilerini kaydeder ve düzenler"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "müþteri raporlama"
            Object.ToolTipText     =   "aranan müþteri gruplarýnýn raporlamasýný yapar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "hakkýnda"
            Object.ToolTipText     =   "ben kimim"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5520
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   255
      _Version        =   393216
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If (Button.Index = 1) Then
Form1.Show
Toolbar1.Buttons(1).Enabled = False
End If

If (Button.Index = 2) Then
Form2.Show
Toolbar1.Buttons(2).Enabled = False
End If



End Sub
