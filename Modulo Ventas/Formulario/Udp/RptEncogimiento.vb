Public Class RptEncogimiento
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ReporteEncogimiento.TextBox1.Text = TextBox3.Text.ToString().Trim()
        ReporteEncogimiento.TextBox2.Text = TextBox2.Text.ToString().Trim()
        ReporteEncogimiento.TextBox3.Text = Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "")
        ReporteEncogimiento.TextBox4.Text = Replace(DateTimePicker2.Value.ToString("yyyy-MM-dd"), "-", "")
        ReporteEncogimiento.ShowDialog()
    End Sub
End Class