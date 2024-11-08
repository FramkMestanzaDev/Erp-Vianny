Public Class Planilla_paqueteoo
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load

        Dim objreporte As New Rpt_planilla_Paqueteo
            objreporte.Refresh()
            objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
            objreporte.SetParameterValue("@CCia", Trim(TextBox1.Text))
            objreporte.SetParameterValue("@Op", Trim(TextBox2.Text))
            If Trim(TextBox3.Text).Length = 0 Then
                objreporte.SetParameterValue("@OCorte", DBNull.Value)
            Else
                objreporte.SetParameterValue("@OCorte", TextBox3.Text)
            End If
            objreporte.SetParameterValue("@Item", Trim(TextBox4.Text))
        objreporte.SetParameterValue("@VerResumen", "1")
        CrystalReportViewer1.ReportSource = objreporte


    End Sub
End Class