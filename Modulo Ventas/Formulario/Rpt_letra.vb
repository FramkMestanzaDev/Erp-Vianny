Public Class Rpt_letra
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Reporte_Letra.TextBox1.Text = Trim(TextBox1.Text)
        Reporte_Letra.TextBox2.Text = Trim(TextBox2.Text)
        Reporte_Letra.Show()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.F1 Then
            Clonar_letra.Label3.Text = 4
            Clonar_letra.Show()

        End If
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.F1 Then
            Clonar_letra.Label3.Text = 5
            Clonar_letra.Show()

        End If
    End Sub
End Class