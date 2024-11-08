Public Class pTERMINADO
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New Rpt_Pt1
        objreporte.Refresh()
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@ccia", TextBox1.Text)
        objreporte.SetParameterValue("@cperiodo", TextBox2.Text)
        objreporte.SetParameterValue("@almacen", TextBox3.Text)
        objreporte.SetParameterValue("@is", TextBox4.Text)
        objreporte.SetParameterValue("@registro", TextBox5.Text)

        CrystalReportViewer1.ReportSource = objreporte
    End Sub
End Class