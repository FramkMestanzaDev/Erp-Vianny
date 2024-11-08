Public Class rpt_anuales
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New Reporte_Ventas_Anuales
        objreporte.Refresh()
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.Refresh()
        objreporte.SetParameterValue("@ANO", TextBox1.Text)
        objreporte.SetParameterValue("@ccia", TextBox2.Text)
        CrystalReportViewer1.ReportSource = objreporte

    End Sub
End Class