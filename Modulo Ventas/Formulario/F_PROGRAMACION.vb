Public Class F_PROGRAMACION
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim j As Integer
        j = Convert.ToInt32(Label3.Text)
        'proveedores.DataGridView1.Rows(j + 1).Selected = True

        Dim fg, fg1, fg2, fg3 As Double

        If Label2.Text = 1 Then
            fg = Convert.ToDouble(proveedores.TextBox9.Text)
            fg2 = Convert.ToDouble(proveedores.TextBox13.Text)
            fg1 = fg + proveedores.DataGridView1.Rows(j).Cells("PROVISION").Value
            fg3 = fg2 + proveedores.DataGridView1.Rows(j).Cells("PROVISION").Value
            proveedores.DataGridView1.Rows(Label3.Text).Cells("F_PROGRA").Value = Format(DateTimePicker1.Value.ToShortDateString, "Short Date")
            proveedores.DataGridView1.Rows(j).Cells("SE").Value = 1
            If Label4.Text = "" Then
                proveedores.TextBox9.Text = fg1
            Else
                proveedores.TextBox13.Text = fg3
            End If

        Else
            fg = Convert.ToDouble(Form_cuentarxcobrar.TextBox9.Text)
            fg2 = Convert.ToDouble(Form_cuentarxcobrar.TextBox13.Text)
            fg1 = fg + Form_cuentarxcobrar.DataGridView1.Rows(j).Cells("PROVISION").Value
            fg3 = fg2 + Form_cuentarxcobrar.DataGridView1.Rows(j).Cells("PROVISION").Value
            Form_cuentarxcobrar.DataGridView1.Rows(Label3.Text).Cells("F_PROGRA").Value = Format(DateTimePicker1.Value.ToShortDateString, "Short Date")
            Form_cuentarxcobrar.DataGridView1.Rows(j).Cells("SE").Value = 1
            Form_cuentarxcobrar.TextBox9.Text = fg1
            If Label4.Text = "S/." Then
                Form_cuentarxcobrar.TextBox9.Text = fg1
            Else
                Form_cuentarxcobrar.TextBox13.Text = fg3
            End If
        End If

        Me.Close()
    End Sub
End Class