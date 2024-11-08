Public Class Rpt_Rama
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New CrystalReport2
        objreporte.Refresh()
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@codigo", TextBox1.Text)

        CrystalReportViewer1.ReportSource = objreporte
    End Sub
End Class