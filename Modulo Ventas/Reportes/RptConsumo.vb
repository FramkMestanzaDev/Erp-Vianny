Public Class RptConsumo
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New CrystalReport4
        objreporte.Refresh()
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        If TextBox1.Text.ToString().Trim() = "" Then
            objreporte.SetParameterValue("@po", DBNull.Value)
        Else
            objreporte.SetParameterValue("@po", TextBox1.Text.ToString().Trim())
        End If

        objreporte.SetParameterValue("@empresa", TextBox2.Text)
        If TextBox3.Text.ToString().Trim() = "" Then
            objreporte.SetParameterValue("@od", DBNull.Value)
        Else
            objreporte.SetParameterValue("@od", TextBox3.Text.ToString().Trim())
        End If
        If TextBox4.Text.ToString().Trim() = "" Then
            objreporte.SetParameterValue("@op", DBNull.Value)
        Else
            objreporte.SetParameterValue("@op", TextBox4.Text.ToString().Trim())
        End If



        CrystalReportViewer1.ReportSource = objreporte

    End Sub
End Class