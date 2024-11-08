Public Class Rpt_Produccion_Corte_Graus
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New Rpt_Prod_Diaria_Corte_Graus
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@finicial", TextBox1.Text)
        objreporte.SetParameterValue("@ffinal", TextBox2.Text)
        CrystalReportViewer1.ReportSource = objreporte
    End Sub
End Class