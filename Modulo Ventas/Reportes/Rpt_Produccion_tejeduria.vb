Public Class Rpt_Produccion_tejeduria
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New roll_teje

        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@ccia", TextBox1.Text)
        objreporte.SetParameterValue("@FECHAINI", Trim(TextBox2.Text))
        objreporte.SetParameterValue("@FECHAFIN", Trim(TextBox3.Text))
        If TextBox4.Text = "" Then
            objreporte.SetParameterValue("@TEJEDOR", DBNull.Value)
        Else
            objreporte.SetParameterValue("@TEJEDOR", TextBox4.Text)
        End If
        If TextBox5.Text = "" Then
            objreporte.SetParameterValue("@MAQUINA", DBNull.Value)
        Else
            objreporte.SetParameterValue("@MAQUINA", TextBox5.Text)
        End If
        objreporte.SetParameterValue("@TIPO_TELA", DBNull.Value)
        objreporte.SetParameterValue("@TITULO", DBNull.Value)
        objreporte.SetParameterValue("@TURNO", DBNull.Value)
        objreporte.SetParameterValue("@NQUIEBRE", 1)
        objreporte.SetParameterValue("@NMODALIDAD", 1)
        objreporte.SetParameterValue("@NAGRUPADO", 1)
        objreporte.SetParameterValue("@CLIENTE", DBNull.Value)
        objreporte.SetParameterValue("@ORDTEJ", DBNull.Value)
        If TextBox6.Text = "" Then
            objreporte.SetParameterValue("@PEDIDOINI", DBNull.Value)
        Else
            objreporte.SetParameterValue("@PEDIDOINI", TextBox6.Text)
        End If
        If TextBox6.Text = "" Then
            objreporte.SetParameterValue("@PEDIDOFIN", DBNull.Value)
        Else
            objreporte.SetParameterValue("@PEDIDOFIN", TextBox6.Text)
        End If

        objreporte.SetParameterValue("@CODTEJEDURIA", DBNull.Value)
        objreporte.SetParameterValue("@LOTE", DBNull.Value)
        CrystalReportViewer1.ReportSource = objreporte
        CrystalReportViewer1.Refresh()
    End Sub
End Class