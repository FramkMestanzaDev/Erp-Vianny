Public Class TEJIDO
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New Rpt_Rollos25
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@PARTIDA", TextBox1.Text)
        objreporte.SetParameterValue("@rollo1", TextBox2.Text)
        objreporte.SetParameterValue("@rollo2", TextBox3.Text)
        objreporte.SetParameterValue("@PEDIDO", TextBox4.Text)
        CrystalReportViewer1.ReportSource = objreporte
    End Sub
End Class