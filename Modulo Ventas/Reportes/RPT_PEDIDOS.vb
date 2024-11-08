Public Class RPT_PEDIDOS
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New PEDIDOSSS

        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@FECHAINI", Trim(TextBox1.Text))
        objreporte.SetParameterValue("@FECHAFIN", Trim(TextBox2.Text))
        CrystalReportViewer1.ReportSource = objreporte
    End Sub
End Class