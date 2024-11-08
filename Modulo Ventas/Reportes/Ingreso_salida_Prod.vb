Public Class Ingreso_salida_Prod
    Private Sub Ingreso_salida_Prod_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim objreporte As New Rpt_ing_sal_prod
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@Ccia", Trim(TextBox1.Text))
        objreporte.SetParameterValue("@Almacen", Trim(TextBox2.Text))
        objreporte.SetParameterValue("@CCom", Trim(TextBox3.Text))
        objreporte.SetParameterValue("@FECHAINI", Trim(TextBox4.Text))
        objreporte.SetParameterValue("@FECHAFIN", Trim(TextBox5.Text))
        If Trim(TextBox6.Text) = "" Then
            objreporte.SetParameterValue("@PEDIDO", DBNull.Value)
        Else
            objreporte.SetParameterValue("@PEDIDO", Trim(TextBox6.Text))
        End If

        objreporte.SetParameterValue("@MOTIVOS", DBNull.Value)
        objreporte.SetParameterValue("@CLIENTE", DBNull.Value)
        objreporte.SetParameterValue("@Origen", "INTEXT")
        objreporte.SetParameterValue("@SubFase", DBNull.Value)
        objreporte.SetParameterValue("@QUIEBRE", 6)
        objreporte.SetParameterValue("@ORDEN", 2)
        CrystalReportViewer1.ReportSource = objreporte
    End Sub

    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load

    End Sub
End Class