﻿Public Class Packing_nacional
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New NACIONAL
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@partida", TextBox1.Text)
        objreporte.SetParameterValue("@calidad", TextBox2.Text)
        objreporte.SetParameterValue("@almacen", TextBox3.Text)
        CrystalReportViewer1.ReportSource = objreporte
    End Sub
End Class