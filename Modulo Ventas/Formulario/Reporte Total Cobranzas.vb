Public Class Reporte_Total_Cobranzas
    Dim dt As DataTable
    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Dim yh As New festadocuenta

        dt = yh.buscar_estado_cuenta2()
        DataGridView1.DataSource = dt
        Dim jk As New Exportar
        jk.llenarExcel(DataGridView1)
    End Sub
End Class