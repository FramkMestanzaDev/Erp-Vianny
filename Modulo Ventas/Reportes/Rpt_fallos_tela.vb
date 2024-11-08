Public Class Rpt_fallos_tela
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New Rpt_calidad_re
        objreporte.Refresh()
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@FECHAINI", TextBox1.Text)
        objreporte.SetParameterValue("@FECHAFINAL", TextBox2.Text)
        objreporte.SetParameterValue("@ETIQUETA", TextBox3.Text)
        CrystalReportViewer1.ReportSource = objreporte
    End Sub
End Class