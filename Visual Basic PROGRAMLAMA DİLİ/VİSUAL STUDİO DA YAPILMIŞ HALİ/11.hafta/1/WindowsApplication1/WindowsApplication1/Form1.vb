Public Class Form1
    Dim isim(3) As String
    Dim notlar(4) As Integer


    Private Sub Form1_Click(sender As Object, e As EventArgs) Handles MyBase.Click
        '' isim(0) = "Ahmet Kagan"
        'isim(1) = "Furkan Zeki"
        'isim(2) = "Adam Otu"
        'notlar(0) = 25
        'notlar(1) = 45
        'notlar(2) = 35
        'notlar(3) = 15
        'Label1.Text = isim(0) + notlar(0)
        'Label2.Text = isim(1) + notlar(1)
        'Label3.Text = isim(2) + notlar(2)
        'Label4.Text = "Görsel Programlama" + notlar(3)
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        isim(0) = "Ahmet Kagan"
        isim(1) = "Furkan Zeki"
        isim(2) = "Adam Otu"
        notlar(0) = 25
        notlar(1) = 45
        notlar(2) = 35
        notlar(3) = 15
        Print(isim(0), notlar(0))
        Print(isim(1), notlar(1))
        Print(isim(2), notlar(2))
        ' Print "Görsel Programlama"; notlar(3)
    End Sub
End Class
