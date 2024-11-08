Public Class Rpte_Partida
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New parti_report
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@partida", TextBox1.Text)
        CrystalReportViewer1.ReportSource = objreporte
    End Sub
End Class