Public Class Rpt_Flujo_Caja_Fecha
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Rpt_cja.TextBox3.Text = Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "")
        Rpt_cja.TextBox4.Text = Replace(DateTimePicker2.Value.ToString("yyyy-MM-dd"), "-", "")
        Rpt_cja.TextBox1.Text = Trim(Label5.Text)
        Rpt_cja.TextBox2.Text = Trim(TextBox2.Text)
        Rpt_cja.Show()
    End Sub
End Class