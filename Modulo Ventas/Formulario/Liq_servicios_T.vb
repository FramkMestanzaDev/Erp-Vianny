Public Class Liq_servicios_T
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.F1 Then
            Lista_Partidas_OT.Label1.Text = Label3.Text
            Lista_Partidas_OT.ShowDialog()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        rrPT_TEJE.TextBox1.Text = Trim(TextBox1.Text)
        rrPT_TEJE.TextBox2.Text = "01"
        rrPT_TEJE.Show()
    End Sub
End Class