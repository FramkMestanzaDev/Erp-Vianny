Public Class Reporte_Consmmo


    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New Consumo_Manufactura
        objreporte.Refresh()
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@ccia", TextBox1.Text)
        objreporte.SetParameterValue("@po", TextBox2.Text)
        CrystalReportViewer1.ReportSource = objreporte
    End Sub
End Class