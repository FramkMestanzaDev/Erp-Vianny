Public Class Reporte_Linea
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New Acta_Entrega_Equipo
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@linea", TextBox1.Text)
        CrystalReportViewer1.ReportSource = objreporte
    End Sub
End Class