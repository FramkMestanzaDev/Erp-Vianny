Public Class REPORTE_CALIDADCC
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim fecini, fecfin As String
        fecini = Format(DateTimePicker1.Value, "yyyy-MM-dd")
        fecfin = Format(DateTimePicker2.Value, "yyyy-MM-dd")
        Rpt_Crudo.TextBox3.Text = (Replace(fecini, "-", ""))
        Rpt_Crudo.TextBox4.Text = (Replace(fecfin, "-", ""))
        If CheckBox1.Checked = True Then
            Rpt_Crudo.TextBox1.Text = TextBox1.Text
        Else
            Rpt_Crudo.TextBox1.Text = ""
        End If
        If CheckBox2.Checked = True Then
            Rpt_Crudo.TextBox2.Text = TextBox2.Text
        Else
            Rpt_Crudo.TextBox2.Text = ""
        End If
        Rpt_Crudo.Show()
    End Sub
    Dim gh As New DataTable
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim JH As New vtejeduria
        Dim LL As New ftejeduria
        DataGridView1.DataSource = ""
        JH.gccia = "01"
        JH.gfechaini = DateTimePicker1.Value
        JH.gfechafin = DateTimePicker2.Value
        If CheckBox1.Checked = True Then
            JH.gpartida = TextBox1.Text
        Else
            JH.gpartida = ""
        End If

        If CheckBox2.Checked = True Then
            JH.gpo = TextBox2.Text
        Else
            JH.gpo = ""
        End If



        gh = LL.reporte_calidad(JH)
        DataGridView1.DataSource = GH
        Dim k As New Exportar
        k.llenarExcel(DataGridView1)
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            TextBox1.Enabled = True
        Else
            TextBox1.Enabled = False
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            TextBox2.Enabled = True
        Else
            TextBox2.Enabled = False
        End If
    End Sub
End Class