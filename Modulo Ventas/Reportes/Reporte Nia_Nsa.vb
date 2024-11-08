Public Class Reporte_Nia_Nsa
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New Nia_Nsa
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@almacen", TextBox1.Text)
        objreporte.SetParameterValue("@condi", TextBox2.Text)
        objreporte.SetParameterValue("@comprobante", TextBox3.Text)
        objreporte.SetParameterValue("@periodo", TextBox4.Text)
        objreporte.SetParameterValue("@ccia", TextBox5.Text)
        CrystalReportViewer1.ReportSource = objreporte
    End Sub

    Private Sub Reporte_Nia_Nsa_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class