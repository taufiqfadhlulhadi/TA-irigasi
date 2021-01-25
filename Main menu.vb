Public Class Main_menu
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim hulu As New Hulu
        hulu.Show()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim tengah As New Tengah
        tengah.Show()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim hilir As New Hilir
        hilir.Show()
    End Sub
End Class