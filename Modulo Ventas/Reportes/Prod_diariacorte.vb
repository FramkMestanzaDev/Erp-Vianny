Public Class Prod_diariacorte
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New Rpt_produccion_diaria_Corte
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@fechaini", TextBox1.Text)
        objreporte.SetParameterValue("@fechafin", TextBox2.Text)
        objreporte.SetParameterValue("@empresa", TextBox3.Text)
        If Trim(TextBox4.Text).Length = 0 Then
            objreporte.SetParameterValue("@op", DBNull.Value)
        Else
            objreporte.SetParameterValue("@op", TextBox4.Text)
        End If

        CrystalReportViewer1.ReportSource = objreporte
    End Sub
End Class