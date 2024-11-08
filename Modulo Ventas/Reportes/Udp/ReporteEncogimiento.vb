Public Class ReporteEncogimiento
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New Encogimiento
        objreporte.Refresh()
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@periodo", TextBox1.Text)
        objreporte.SetParameterValue("@empresa", TextBox2.Text)
        objreporte.SetParameterValue("@fehaini", TextBox3.Text)
        objreporte.SetParameterValue("@fehafin", TextBox4.Text)
        CrystalReportViewer1.ReportSource = objreporte
    End Sub
End Class