Public Class rPT_PACKIN_HIANCHIPA
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New Rpt_Partidaa
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@partida", TextBox1.Text)
        objreporte.SetParameterValue("@almacen", TextBox2.Text)
        objreporte.SetParameterValue("@AL", TextBox3.Text)
        CrystalReportViewer1.ReportSource = objreporte
    End Sub
End Class