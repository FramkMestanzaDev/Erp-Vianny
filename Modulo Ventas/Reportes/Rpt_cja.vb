Public Class Rpt_cja
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New Rpt_Flujo_Caja
        objreporte.Refresh()
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@ccia", TextBox1.Text)
        objreporte.SetParameterValue("@CPERIODO", TextBox2.Text)
        objreporte.SetParameterValue("@FECHAINI", TextBox3.Text)
        objreporte.SetParameterValue("@FECHAFIN", TextBox4.Text)
        objreporte.SetParameterValue("@CUENTA_INI", "10411101")
        objreporte.SetParameterValue("@CUENTA_FIN", "10412107")
        objreporte.SetParameterValue("@nmoneda", "1")
        CrystalReportViewer1.ReportSource = objreporte
    End Sub
End Class