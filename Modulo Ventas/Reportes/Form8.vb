Public Class Form8
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New Rpt_Factura
        objreporte.Refresh()
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@sfactu_3", TextBox1.Text)
        objreporte.SetParameterValue("@nfactu_3", TextBox2.Text)
        objreporte.SetParameterValue("@tienda_3", TextBox3.Text)
        objreporte.SetParameterValue("@ccia_3", TextBox4.Text)
        CrystalReportViewer1.ReportSource = objreporte
    End Sub
End Class