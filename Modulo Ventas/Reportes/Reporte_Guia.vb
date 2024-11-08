Public Class Reporte_Guia
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New Guia_Remision6
        objreporte.Refresh()
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@sfactu_3a", TextBox1.Text)
        objreporte.SetParameterValue("@nfactu_3a", TextBox2.Text)
        objreporte.SetParameterValue("@ccia", TextBox3.Text)
        objreporte.SetParameterValue("@op", TextBox4.Text)
        CrystalReportViewer1.ReportSource = objreporte
    End Sub
End Class