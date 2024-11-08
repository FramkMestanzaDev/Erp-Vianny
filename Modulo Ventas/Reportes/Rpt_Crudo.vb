Public Class Rpt_Crudo
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New crudo
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        If Trim(TextBox1.Text) = "" Then
            objreporte.SetParameterValue("@PARTIDA", DBNull.Value)
        Else
            objreporte.SetParameterValue("@PARTIDA", TextBox1.Text)
        End If
        If Trim(TextBox2.Text) = "" Then
            objreporte.SetParameterValue("@PEDIDO", DBNull.Value)
        Else
            objreporte.SetParameterValue("@PEDIDO", TextBox2.Text)
        End If

        objreporte.SetParameterValue("@FECHAINI", TextBox3.Text)
        objreporte.SetParameterValue("@FECHAFIN", TextBox4.Text)
        CrystalReportViewer1.ReportSource = objreporte
    End Sub
End Class