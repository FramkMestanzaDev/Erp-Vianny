Public Class Rpt_ventas_rubro
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New reporte_rubro
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@ANO", TextBox1.Text)
        objreporte.SetParameterValue("@ccia", TextBox2.Text)
        If TextBox3.Text = "NULL" Then
            objreporte.SetParameterValue("@mes", DBNull.Value)
        Else
            objreporte.SetParameterValue("@mes", Trim(TextBox3.Text))
        End If

        CrystalReportViewer1.ReportSource = objreporte
    End Sub
End Class