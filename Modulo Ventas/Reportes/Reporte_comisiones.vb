Public Class Reporte_comisiones
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        'Dim objreporte As New Comisiones2
        'objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%", "172.21.0.1", "Comercial_Textil")
        'objreporte.SetParameterValue("@vendedor", TextBox1.Text)
        'objreporte.SetParameterValue("@mes", TextBox2.Text)
        'objreporte.SetParameterValue("@PERIODO", TextBox3.Text)
        'objreporte.SetParameterValue("@ccia", TextBox4.Text)
        'CrystalReportViewer1.ReportSource = objreporte
        Dim objreporte As New comision_vendedor
        objreporte.Refresh()
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%", "172.21.0.1", "Comercial_Textil")
        objreporte.SetParameterValue("@id", TextBox1.Text)


        CrystalReportViewer1.ReportSource = objreporte
    End Sub


End Class