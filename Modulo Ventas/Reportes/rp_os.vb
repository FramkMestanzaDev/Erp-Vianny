Public Class rp_os
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New Rpt_servicio_acumulado
        objreporte.Refresh()
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@fini", TextBox1.Text)
        objreporte.SetParameterValue("@ffin", TextBox2.Text)
        objreporte.SetParameterValue("@ccia", TextBox3.Text)
        If Trim(TextBox4.Text).Length = 0 Then
            objreporte.SetParameterValue("@op", DBNull.Value)
        Else
            objreporte.SetParameterValue("@op", TextBox4.Text)
        End If
        If Trim(TextBox5.Text).Length = 0 Then
            objreporte.SetParameterValue("@ruc", DBNull.Value)
        Else
            objreporte.SetParameterValue("@ruc", TextBox5.Text)
        End If

        CrystalReportViewer1.ReportSource = objreporte
    End Sub
End Class