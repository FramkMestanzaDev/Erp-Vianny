Public Class Rpt_Tejeduria
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New OT
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@ccia", "01")
        objreporte.SetParameterValue("@orden_tejido", TextBox1.Text)
        objreporte.SetParameterValue("@POINI", DBNull.Value)
        objreporte.SetParameterValue("@POFIN", DBNull.Value)
        objreporte.SetParameterValue("@rangopo", DBNull.Value)
        objreporte.SetParameterValue("@cliente", DBNull.Value)
        objreporte.SetParameterValue("@f_emisionini", DBNull.Value)
        objreporte.SetParameterValue("@f_emisionfin", DBNull.Value)
        CrystalReportViewer1.ReportSource = objreporte
    End Sub
End Class