Public Class Rpt_Calidad_acabado


    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        ' Dim objreporte As New Rpt_calidad_re
        Dim objreporte As New Rpt_fallas
        objreporte.Refresh()
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@PROGRAMA", TextBox1.Text)
        CrystalReportViewer1.ReportSource = objreporte
    End Sub
End Class