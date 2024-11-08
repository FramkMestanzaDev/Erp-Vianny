Public Class Rpt_Packing
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New Rpt_Packing_Tela
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@GUIA", TextBox1.Text)

        CrystalReportViewer1.ReportSource = objreporte
    End Sub
End Class