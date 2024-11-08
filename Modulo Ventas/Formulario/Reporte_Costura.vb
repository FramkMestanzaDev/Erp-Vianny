Public Class Reporte_Costura
    Dim dt As New DataTable
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim jh As New forden_costura
        Dim k As Integer
        dt = jh.COSTURA_CONTROLER()
        If dt.Rows.Count <> 0 Then
            DataGridView1.DataSource = dt
            k = DataGridView1.Rows.Count
            For i = 0 To k - 1
                If DataGridView1.Rows(i).Cells(4).Value IsNot DBNull.Value Then
                    DataGridView1.Rows(i).Cells(4).Value = "ASIGNADO"
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.DarkRed
                    DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.White
                Else
                    DataGridView1.Rows(i).Cells(4).Value = "POR ASIGNAR"
                End If
            Next

        Else
            MsgBox("NO EXISTE INFORMION PARA MOSTRAR")
        End If
    End Sub
End Class