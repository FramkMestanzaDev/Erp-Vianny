Public Class Rpt_cuentaxPagar
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If RadioButton1.Checked = True Then
            Rpt_proveedores.TextBox1.Text = TextBox1.Text
            Rpt_proveedores.TextBox2.Text = Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "")
            Rpt_proveedores.TextBox3.Text = Replace(DateTimePicker2.Value.ToString("yyyy-MM-dd"), "-", "")
            Rpt_proveedores.TextBox4.Text = Label5.Text
            Rpt_proveedores.Show()
        Else
            ACUMULADO_PROVEEDORES.TextBox1.Text = TextBox1.Text
            ACUMULADO_PROVEEDORES.TextBox2.Text = Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "")
            ACUMULADO_PROVEEDORES.TextBox3.Text = Replace(DateTimePicker2.Value.ToString("yyyy-MM-dd"), "-", "")
            ACUMULADO_PROVEEDORES.TextBox4.Text = Label5.Text
            ACUMULADO_PROVEEDORES.Show()
        End If

    End Sub

    Private Sub Rpt_cuentaxPagar_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        RadioButton1.Checked = True
    End Sub
End Class