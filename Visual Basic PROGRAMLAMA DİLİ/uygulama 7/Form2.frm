VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Form2"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3135
   ScaleWidth      =   4680
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Unload(Cancel As Integer)
MDIForm1.Toolbar1.Buttons(2).Enabled = True
End Sub
