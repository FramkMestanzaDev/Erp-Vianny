Public Class ReporteMatriz
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New RptMatrizAvios
        objreporte.Refresh()
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@CCIA", Trim(TextBox1.Text))
        objreporte.SetParameterValue("@PEDIDO_INI", Trim(TextBox2.Text))
        objreporte.SetParameterValue("@PEDIDO_FIN", Trim(TextBox2.Text))
        objreporte.SetParameterValue("@RELACION_PEDIDOS", DBNull.Value)
        objreporte.SetParameterValue("@AREA", "01")
        CrystalReportViewer1.ReportSource = objreporte
    End Sub
End Class