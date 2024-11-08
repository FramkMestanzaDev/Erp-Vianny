Public Class Imp_Expotracion
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New imprimi_rollos
        objreporte.Refresh()
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@partida", TextBox1.Text)
        objreporte.SetParameterValue("@rollo", TextBox2.Text)
        objreporte.SetParameterValue("@pesoa", TextBox3.Text)
        objreporte.SetParameterValue("@calidad", TextBox4.Text)
        objreporte.SetParameterValue("@ancho", TextBox5.Text)
        objreporte.SetParameterValue("@orden", TextBox6.Text)
        objreporte.SetParameterValue("@densidad", TextBox7.Text)
        CrystalReportViewer1.ReportSource = objreporte
    End Sub
End Class