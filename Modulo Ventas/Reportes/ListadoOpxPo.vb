Public Class ListadoOpxPo
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New RptListadoPOOP
        objreporte.Refresh()
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@ccIA", TextBox1.Text)
        objreporte.SetParameterValue("@PO", TextBox2.Text)
        CrystalReportViewer1.ReportSource = objreporte
    End Sub
End Class