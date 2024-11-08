Public Class Rppt_guiar
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New Guia_Remision222
        objreporte.Refresh()
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@CCia", Trim(TextBox1.Text))
        objreporte.SetParameterValue("@nTipo", 1)
        objreporte.SetParameterValue("@Fecha_Ini", Trim(TextBox3.Text))
        objreporte.SetParameterValue("@Fecha_Fin", Trim(TextBox4.Text))
        objreporte.SetParameterValue("@Serie_Ini", DBNull.Value)
        objreporte.SetParameterValue("@Numero_Ini", DBNull.Value)
        objreporte.SetParameterValue("@Serie_Fin", DBNull.Value)
        objreporte.SetParameterValue("@Numero_Fin", DBNull.Value)

        If TextBox9.Text = "0" Then
            objreporte.SetParameterValue("@Ruc", DBNull.Value)
        Else
            objreporte.SetParameterValue("@Ruc", TextBox9.Text)
        End If
        If TextBox10.Text = "0" Then
            objreporte.SetParameterValue("@Serie", DBNull.Value)
        Else
            objreporte.SetParameterValue("@Serie", TextBox10.Text)
        End If
        If TextBox11.Text = "0" Then
            objreporte.SetParameterValue("@MotivoTraslado", DBNull.Value)
        Else
            objreporte.SetParameterValue("@MotivoTraslado", TextBox11.Text)
        End If

        objreporte.SetParameterValue("@nSituacion", 1)
        objreporte.SetParameterValue("@nQuiebre", 1)
        objreporte.SetParameterValue("@nOrden", 1)
        CrystalReportViewer1.ReportSource = objreporte
    End Sub
End Class