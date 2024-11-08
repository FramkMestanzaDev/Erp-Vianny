Public Class Rpt_frecuencia_ventas
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New Rpt_frecventas
        objreporte.Refresh()
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@ano", TextBox1.Text)
        objreporte.SetParameterValue("@empresa", TextBox2.Text)
        objreporte.SetParameterValue("@vendedor", TextBox3.Text)
        objreporte.SetParameterValue("@nvendedor", TextBox4.Text)
        CrystalReportViewer1.ReportSource = objreporte
    End Sub
End Class