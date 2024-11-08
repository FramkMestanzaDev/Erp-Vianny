Public Class Consumo_tela_Consolidado
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.F1 Then
            DET_FICHA_TECNICA.Close()
            DET_FICHA_TECNICA.Label3.Text = 8
            DET_FICHA_TECNICA.ShowDialog()
        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Reporte_Consmmo.TextBox1.Text = "01"
        Reporte_Consmmo.TextBox2.Text = TextBox1.Text
        Reporte_Consmmo.Show()
    End Sub
End Class