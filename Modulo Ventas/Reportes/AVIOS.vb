Public Class AVIOS
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New Requerimiento_avios_detalle
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@id_requeminieto", TextBox1.Text)
        objreporte.SetParameterValue("@empresa", TextBox2.Text)
        objreporte.SetParameterValue("@periodo", TextBox3.Text)
        objreporte.SetParameterValue("@detalle", TextBox4.Text)
        CrystalReportViewer1.ReportSource = objreporte
    End Sub


End Class