Public Class Rpt_oc_1
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New Rpt_oc
        objreporte.Refresh()
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@CCIA", TextBox1.Text)
        objreporte.SetParameterValue("@ALMACENI", TextBox2.Text)
        objreporte.SetParameterValue("@NRO_ORDENI", TextBox3.Text)
        objreporte.SetParameterValue("@ALMACENf", TextBox4.Text)
        objreporte.SetParameterValue("@NRO_ORDENf", TextBox5.Text)
        objreporte.SetParameterValue("@TIPO_FORMATO", DBNull.Value)
        objreporte.SetParameterValue("@ver_Cant_Envio", TextBox7.Text)
        objreporte.SetParameterValue("@RutaImagen", TextBox8.Text)
        CrystalReportViewer1.ReportSource = objreporte
    End Sub
End Class