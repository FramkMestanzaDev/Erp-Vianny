Public Class status
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New Status_Produccion
        objreporte.Refresh()
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@po", Trim(TextBox1.Text))
        objreporte.SetParameterValue("@ccia", Trim(TextBox2.Text))

        CrystalReportViewer1.ReportSource = objreporte
    End Sub
End Class