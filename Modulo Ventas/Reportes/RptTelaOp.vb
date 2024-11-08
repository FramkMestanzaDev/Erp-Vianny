Public Class RptTelaOp
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New ReporteTelaOp
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@empresa", TextBox1.Text)
        objreporte.SetParameterValue("@op", TextBox2.Text)
        CrystalReportViewer1.ReportSource = objreporte
    End Sub
End Class