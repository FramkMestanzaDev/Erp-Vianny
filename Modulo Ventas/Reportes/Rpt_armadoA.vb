Public Class Rpt_armadoA
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New Rpt_armado_avios
        objreporte.Refresh()
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@id_requeminieto", Trim(TextBox1.Text))
        objreporte.SetParameterValue("@empresa", Trim(TextBox2.Text))
        objreporte.SetParameterValue("@periodo", Trim(TextBox3.Text))
        objreporte.SetParameterValue("@parcial", Trim(TextBox4.Text))
        CrystalReportViewer1.ReportSource = objreporte
    End Sub
End Class