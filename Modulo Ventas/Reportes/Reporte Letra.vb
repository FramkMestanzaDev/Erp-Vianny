Public Class Reporte_Letra


    Private Sub CrystalReportViewer1_Load_1(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New RptLetra
        objreporte.Refresh()
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@LetIni", Trim(TextBox1.Text))
        objreporte.SetParameterValue("@LetFin", Trim(TextBox2.Text))
        CrystalReportViewer1.ReportSource = objreporte
    End Sub


End Class