Public Class Imprimir_Teñido
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Report_Tenido.TextBox1.Text = TextBox1.Text
        Report_Tenido.Show()
    End Sub
End Class