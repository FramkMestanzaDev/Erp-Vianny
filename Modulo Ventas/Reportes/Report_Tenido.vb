﻿Public Class Report_Tenido
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New Rpt_tenido
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@tenido", TextBox1.Text)

        CrystalReportViewer1.ReportSource = objreporte
    End Sub
End Class