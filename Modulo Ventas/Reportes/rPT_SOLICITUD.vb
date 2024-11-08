Public Class rPT_SOLICITUD
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New Rpt_requerimiento_avios
        objreporte.Refresh()
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@id_requeminieto", TextBox1.Text)
        objreporte.SetParameterValue("@empresa", TextBox2.Text)
        objreporte.SetParameterValue("@periodo", TextBox3.Text)
        CrystalReportViewer1.ReportSource = objreporte
    End Sub
End Class