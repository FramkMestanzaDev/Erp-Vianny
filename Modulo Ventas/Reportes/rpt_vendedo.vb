Public Class rpt_vendedo
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New Report_ventas_vendedor
        objreporte.Refresh()
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")

        objreporte.SetParameterValue("@ANO", TextBox1.Text)
        objreporte.SetParameterValue("@ccia", TextBox2.Text)
        CrystalReportViewer1.ReportSource = objreporte

    End Sub
End Class