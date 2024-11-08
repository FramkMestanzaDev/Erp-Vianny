Public Class ACUMULADO_PROVEEDORES
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New XPAGARACUMULADO
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@periodo", TextBox1.Text)
        objreporte.SetParameterValue("@fini", TextBox2.Text)
        objreporte.SetParameterValue("@ffin", TextBox3.Text)
        objreporte.SetParameterValue("@ccia", TextBox4.Text)
        CrystalReportViewer1.ReportSource = objreporte
    End Sub
End Class