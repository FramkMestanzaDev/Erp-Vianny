Public Class Rpt_Explocion
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New CrystalReport3
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@ccia", Trim(TextBox2.Text))
        objreporte.SetParameterValue("@op", Trim(TextBox1.Text))
        CrystalReportViewer1.ReportSource = objreporte
    End Sub
End Class