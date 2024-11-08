Public Class REQ_C
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New req_comercial
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@numero", TextBox1.Text)

        CrystalReportViewer1.ReportSource = objreporte
    End Sub
End Class