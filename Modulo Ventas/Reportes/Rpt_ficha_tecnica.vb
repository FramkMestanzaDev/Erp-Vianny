Public Class Rpt_ficha_tecnica
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New Rpt_ficha__tecnica
        objreporte.Refresh()
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@ccia", Trim(TextBox2.Text))
        objreporte.SetParameterValue("@CODIGO_TELA", Trim(TextBox1.Text))
        objreporte.SetParameterValue("@NORDEN", 1)

        CrystalReportViewer1.ReportSource = objreporte
    End Sub
End Class