Public Class Rpt_export
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New imprimi_rollos_exportacion
        objreporte.Refresh()
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@partida", TextBox1.Text)
        objreporte.SetParameterValue("@orden", TextBox2.Text)
        CrystalReportViewer1.ReportSource = objreporte
    End Sub
End Class