﻿Public Class ReporteMatrizPo
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New RptMatrizPo
        objreporte.Refresh()
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@CCIA", Trim(TextBox1.Text))
        objreporte.SetParameterValue("@PO", Trim(TextBox2.Text))
        objreporte.SetParameterValue("@AREA", "01")
        CrystalReportViewer1.ReportSource = objreporte
    End Sub
End Class