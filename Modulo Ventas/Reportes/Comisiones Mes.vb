Public Class Comisiones_Mes
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New Rpt_comisiones3
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@mes", TextBox1.Text)
        objreporte.SetParameterValue("@periodo", TextBox2.Text)
        objreporte.SetParameterValue("@ccia", TextBox3.Text)
        CrystalReportViewer1.ReportSource = objreporte
    End Sub
End Class