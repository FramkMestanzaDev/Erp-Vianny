﻿Public Class Rpt_Pa_Rama
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        Dim objreporte As New Rpt_prama
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@codigo", TextBox1.Text)

        CrystalReportViewer1.ReportSource = objreporte
    End Sub
End Class