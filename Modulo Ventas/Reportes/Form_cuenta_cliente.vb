Public Class Form_cuenta_cliente
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New Estado_Cuenta
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@fecha", TextBox1.Text)
        objreporte.SetParameterValue("@ruc", TextBox2.Text)
        objreporte.SetParameterValue("@ccia", TextBox3.Text)
        CrystalReportViewer1.ReportSource = objreporte
    End Sub
End Class