Public Class Prod_corteeee
    Dim fecha1, fecha2 As String

    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        fecha1 = Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "")
        fecha2 = Replace(DateTimePicker2.Value.ToString("yyyy-MM-dd"), "-", "")


        Dim objreporte As New Rpt_corte_graus
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@FECHAFIN", fecha2)
        objreporte.SetParameterValue("@FECHAINI", fecha1)
        CrystalReportViewer1.ReportSource = objreporte

    End Sub
End Class