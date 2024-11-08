Public Class ReportePm
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        If Trim(TextBox3.Text) = "1" Then
            Dim objreporte As New RptListadoPMOD
            objreporte.Refresh()
            objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
            objreporte.SetParameterValue("@ccIA", TextBox1.Text)
            objreporte.SetParameterValue("@PO", TextBox2.Text)
            CrystalReportViewer1.ReportSource = objreporte
        Else
            Dim objreporte As New RptListadoPOOP
            objreporte.Refresh()
            objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
            objreporte.SetParameterValue("@ccIA", TextBox1.Text)
            objreporte.SetParameterValue("@PO", TextBox2.Text)
            CrystalReportViewer1.ReportSource = objreporte
        End If

    End Sub
End Class