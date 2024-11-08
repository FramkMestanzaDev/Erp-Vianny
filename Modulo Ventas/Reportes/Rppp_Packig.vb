Public Class Rppp_Packig
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New paaacking
        objreporte.Refresh()
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@ccia", "01")
        objreporte.SetParameterValue("@pedidoot", TextBox1.Text)
        objreporte.SetParameterValue("@correlativo", TextBox2.Text)
        objreporte.SetParameterValue("@ordetejido", DBNull.Value)
        objreporte.SetParameterValue("@nrorolloini", DBNull.Value)
        objreporte.SetParameterValue("@nrorollofin", DBNull.Value)

        objreporte.SetParameterValue("@lista_rollos", TextBox3.Text)

        objreporte.SetParameterValue("@pesoa", TextBox4.Text)
        If TextBox5.Text = "" Then
            objreporte.SetParameterValue("@calidad", DBNull.Value)
        Else
            objreporte.SetParameterValue("@calidad", TextBox5.Text)
        End If
        objreporte.SetParameterValue("@ancho", TextBox6.Text)
        CrystalReportViewer1.ReportSource = objreporte
        'With CrystalReportViewer1
        '    objreporte.PrintToPrinter(1, False, 0, 0)
        'End With
    End Sub
End Class