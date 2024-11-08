Public Class rrPT_TEJE
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New Rpt_Liq_S_T
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@PARTIDA", TextBox1.Text)
        objreporte.SetParameterValue("@CCIA", "01")
        CrystalReportViewer1.ReportSource = objreporte
    End Sub
End Class