Public Class FORM_OS
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New Rpt_os
        objreporte.Refresh()
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@orden", Trim(TextBox1.Text))
        objreporte.SetParameterValue("@orden2", Trim(TextBox2.Text))
        objreporte.SetParameterValue("@ccia", Trim(TextBox3.Text))
        objreporte.SetParameterValue("@RutaImagen", Trim(TextBox4.Text))
        CrystalReportViewer1.ReportSource = objreporte
    End Sub
End Class