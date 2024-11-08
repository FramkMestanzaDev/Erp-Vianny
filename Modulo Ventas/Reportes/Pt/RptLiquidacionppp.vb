Public Class RptLiquidacionppp
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New Rpt_Liquidacion()
        objreporte.Refresh()
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@Mes", TextBox1.Text)
        objreporte.SetParameterValue("@ccia", TextBox2.Text)
        objreporte.SetParameterValue("@año", TextBox3.Text)
        CrystalReportViewer1.ReportSource = objreporte
    End Sub
End Class