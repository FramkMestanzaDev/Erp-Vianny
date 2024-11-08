Public Class Form_ventan
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New Rpt_venta_libre
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@factu", TextBox1.Text)
        objreporte.SetParameterValue("@nfactu", TextBox2.Text)
        CrystalReportViewer1.ReportSource = objreporte
    End Sub
End Class