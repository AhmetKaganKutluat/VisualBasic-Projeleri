Public Class Form3

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Shell("C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE", vbNormalFocus)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Shell("C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE", vbNormalFocus)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Shell("notepad bilgiler.txt", vbNormalFocus)
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Shell("calc.exe", vbNormalFocus)
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Shell("mspaint.exe", vbNormalFocus)
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Shell("C:\Program Files (x86)\Google\Chrome\Application\chrome.exe www.akdeniz.edu.tr", vbNormalFocus)
    End Sub
    Dim sor As Integer
    '' Dim yedek As New FileSystemObject
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        '' proje menüsünden referanslara geliyoruz . microsoft scripting runtime seçilir.
        '' yukarı bi değişken atarız genelde çalışan
        sor = MsgBox("BİLGİLER.TXT DOSYASININ YEDEĞİNİ ALMAK İSTİYORMUSUNZ ? ", 4 + 32, "KOPYALAMA İŞLEMİ ")
        If sor = 6 Then
            ''  yedek.CopyFile("bilgiler.txt", "yedek.txt", True)
        End If
    End Sub
End Class