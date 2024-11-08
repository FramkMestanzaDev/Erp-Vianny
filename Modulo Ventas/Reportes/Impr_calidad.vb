Public Class Impr_calidad
    Private Sub Impr_calidad_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim objreporte As New paaackingCALIDAD
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@ccia", "01")
        objreporte.SetParameterValue("@pedidoot", TextBox1.Text)
        objreporte.SetParameterValue("@correlativo", TextBox2.Text)
        objreporte.SetParameterValue("@ordetejido", DBNull.Value)
        objreporte.SetParameterValue("@nrorolloini", DBNull.Value)
        objreporte.SetParameterValue("@nrorollofin", DBNull.Value)
        objreporte.SetParameterValue("@lista_rollos", DBNull.Value)
        objreporte.SetParameterValue("@pesoa", DBNull.Value)
        objreporte.SetParameterValue("@calidad", DBNull.Value)
        objreporte.SetParameterValue("@ancho", DBNull.Value)
        CrystalReportViewer1.ReportSource = objreporte
    End Sub

    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load

    End Sub
End Class