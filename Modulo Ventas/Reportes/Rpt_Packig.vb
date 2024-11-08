Public Class Rpt_Packig
    Private Sub Rpt_Packig_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim objreporte As New pR_pACKING
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@partida", TextBox1.Text)
        objreporte.SetParameterValue("@calidad", TextBox2.Text)
        CrystalReportViewer1.ReportSource = objreporte
    End Sub

    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load

    End Sub
End Class