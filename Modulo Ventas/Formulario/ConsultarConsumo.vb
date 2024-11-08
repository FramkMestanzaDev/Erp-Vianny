Public Class ConsultarConsumo
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox5.Text.ToString().Trim().Length() = 0 Then
            RptConsumo.TextBox1.Text = TextBox2.Text.ToString().Trim()
            RptConsumo.TextBox2.Text = Label3.Text
            RptConsumo.TextBox3.Text = TextBox4.Text.ToString().Trim()
            RptConsumo.TextBox4.Text = TextBox3.Text.ToString().Trim()
            If CheckBox1.Checked = True Then
                RptConsumo.TextBox5.Text = Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "")
                RptConsumo.TextBox6.Text = Replace(DateTimePicker2.Value.ToString("yyyy-MM-dd"), "-", "")
            Else
                RptConsumo.TextBox5.Text = ""
                RptConsumo.TextBox6.Text = ""
            End If
            RptConsumo.ShowDialog(Me)
        Else
            RptConsumoProducto.TextBox1.Text = Label3.Text
            RptConsumoProducto.TextBox2.Text = TextBox5.Text.ToString().Trim()
            RptConsumoProducto.ShowDialog(Me)
        End If

    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.F1 Then
            ListaPO.Label3.Text = "4"
            ListaPO.ShowDialog(Me)
        End If

    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            TextBox2.Enabled = True
            TextBox3.Enabled = False
            TextBox4.Enabled = False
            TextBox5.Enabled = False
            DateTimePicker1.Enabled = False
            DateTimePicker2.Enabled = False
            CheckBox1.Checked = False
            CheckBox3.Checked = False
            CheckBox4.Checked = False
            CheckBox5.Checked = False
            TextBox2.Select()
        Else
            TextBox2.Text = ""
            TextBox2.Enabled = False
        End If
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = True Then
            TextBox2.Enabled = False
            TextBox3.Enabled = True
            TextBox4.Enabled = False
            TextBox5.Enabled = False
            DateTimePicker1.Enabled = False
            DateTimePicker2.Enabled = False
            CheckBox1.Checked = False
            CheckBox2.Checked = False
            CheckBox4.Checked = False
            CheckBox5.Checked = False
            TextBox3.Select()
        Else
            TextBox3.Text = ""
            TextBox3.Enabled = False
        End If
    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.Checked = True Then
            TextBox2.Enabled = False
            TextBox3.Enabled = False
            TextBox4.Enabled = True
            TextBox5.Enabled = False
            DateTimePicker1.Enabled = False
            DateTimePicker2.Enabled = False
            CheckBox1.Checked = False
            CheckBox2.Checked = False
            CheckBox3.Checked = False
            CheckBox5.Checked = False
            TextBox4.Select()
        Else
            TextBox4.Text = ""
            TextBox4.Enabled = False
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            TextBox2.Enabled = False
            TextBox3.Enabled = False
            TextBox4.Enabled = False
            TextBox5.Enabled = False
            DateTimePicker1.Enabled = True
            DateTimePicker2.Enabled = True
            CheckBox3.Checked = False
            CheckBox2.Checked = False
            CheckBox4.Checked = False
            CheckBox5.Checked = False
        Else

            DateTimePicker1.Enabled = False
            DateTimePicker2.Enabled = False
        End If
    End Sub

    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged
        If CheckBox5.Checked = True Then
            TextBox2.Enabled = False
            TextBox3.Enabled = False
            TextBox4.Enabled = False
            TextBox5.Enabled = True
            DateTimePicker1.Enabled = False
            DateTimePicker2.Enabled = False
            CheckBox1.Checked = False
            CheckBox2.Checked = False
            CheckBox3.Checked = False
            CheckBox4.Checked = False
            TextBox5.Select()
        Else
            TextBox5.Text = ""
            TextBox5.Enabled = False
        End If
    End Sub

    Private Sub TextBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox5.KeyDown
        If e.KeyCode = Keys.F1 Then
            pproductos.CCIA.Text = Label3.Text
            pproductos.ALMACEN.Text = "08"
            pproductos.ANO.Text = MDIParent2.Label8.Text
            pproductos.FECHA.Text = Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "")
            pproductos.TextBox1.Text = Trim(TextBox1.Text)
            pproductos.Label3.Text = 15
            pproductos.Label5.Text = ""
            pproductos.Show()
        End If
    End Sub
End Class