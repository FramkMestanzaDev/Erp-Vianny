Public Class Reporte_ptt
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New Producto_ter
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@ccia", TextBox1.Text)
        objreporte.SetParameterValue("@periodo", TextBox2.Text)

        CrystalReportViewer1.ReportSource = objreporte
    End Sub
End Class