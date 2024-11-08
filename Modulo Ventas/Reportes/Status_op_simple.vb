Public Class Status_op_simple
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New status_op
        objreporte.Refresh()
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@OP", Trim(TextBox1.Text))
        objreporte.SetParameterValue("@ccia", Trim(TextBox2.Text))

        CrystalReportViewer1.ReportSource = objreporte
    End Sub
End Class