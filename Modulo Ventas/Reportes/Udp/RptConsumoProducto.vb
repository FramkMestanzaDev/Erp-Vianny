Public Class RptConsumoProducto
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New RptConsumoTelacodigo
        objreporte.Refresh()
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@empresa", TextBox1.Text)
        objreporte.SetParameterValue("@codigo", TextBox2.Text)
        CrystalReportViewer1.ReportSource = objreporte
    End Sub
End Class