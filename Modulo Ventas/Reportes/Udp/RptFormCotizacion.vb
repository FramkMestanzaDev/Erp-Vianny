Public Class RptFormCotizacion
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New RptCotizacion
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.Refresh()
        objreporte.SetParameterValue("@cotizacion", TextBox1.Text)
        objreporte.SetParameterValue("@AREA", TextBox2.Text)
        CrystalReportViewer1.ReportSource = objreporte
    End Sub
End Class