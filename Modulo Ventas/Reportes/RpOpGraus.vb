Public Class RpOpGraus


    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New ReporteOPGRaus
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@ccia", TextBox1.Text)
        objreporte.SetParameterValue("@PEDIDO_INI", TextBox2.Text)
        objreporte.SetParameterValue("@PEDIDO_FIN", TextBox3.Text)
        objreporte.SetParameterValue("@INC_CODCLI", TextBox4.Text)
        objreporte.SetParameterValue("@VALORIZADO", TextBox5.Text)
        CrystalReportViewer1.ReportSource = objreporte
    End Sub
End Class