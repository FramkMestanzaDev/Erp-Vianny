Public Class det_maquina
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        DataGridView1.Rows.Add(ComboBox1.Text)

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim t As Integer

        If Label2.Text = 1 Then
            MAQUINAS_TEJEDORAS.DataGridView2.Rows.Clear()
            t = DataGridView1.Rows.Count

            MAQUINAS_TEJEDORAS.DataGridView2.Rows.Add(t)

            For i = 0 To t - 1

                MAQUINAS_TEJEDORAS.DataGridView2.Rows(i).Cells(0).Value = DataGridView1.Rows(i).Cells(0).Value
            Next
        Else
            If Label2.Text = "2" Then
                MAQUINAS_TEJEDORAS.DataGridView3.Rows.Clear()
                t = DataGridView1.Rows.Count
                MAQUINAS_TEJEDORAS.DataGridView3.Rows.Add(t)

                For i = 0 To t - 1
                    MAQUINAS_TEJEDORAS.DataGridView3.Rows(i).Cells(0).Value = DataGridView1.Rows(i).Cells(0).Value
                Next
            End If
        End If
        Me.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim respuesta As DialogResult

        respuesta = MessageBox.Show("DESEA ELIMINAR LA FILA ?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then
            DataGridView1.Rows.Remove(DataGridView1.CurrentRow)

        End If
    End Sub

    Private Sub det_maquina_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class