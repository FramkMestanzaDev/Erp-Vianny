Public Class Rpt_corte
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New Reporte_corte
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@op", TextBox1.Text)
        objreporte.SetParameterValue("@corte", TextBox2.Text)
        CrystalReportViewer1.ReportSource = objreporte
    End Sub
End Class