Public Class Rpt_Calidad
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New Reporte_Calidad
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@PO", TextBox1.Text)

        CrystalReportViewer1.ReportSource = objreporte
    End Sub
End Class