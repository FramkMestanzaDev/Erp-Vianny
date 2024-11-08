Public Class Confeccion
    Dim dt As New DataTable
    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            actualizar()
        End If
    End Sub
    Sub actualizar()
        Dim fo As New forden_costura
        Dim fo1 As New vorden_costura
        Dim Ia As Integer
        fo1.gop = TextBox1.Text
        dt = fo.buscar_costura(fo1)
        If dt.Rows.Count <> 0 Then
            DataGridView1.DataSource = dt
            DataGridView1.Columns(1).Width = 220
            DataGridView1.Columns(2).Width = 90
            DataGridView1.Columns(3).Width = 90
            DataGridView1.Columns(4).Width = 70

            Ia = DataGridView1.Rows.Count
            For I = 0 To Ia - 1

                If DataGridView1.Rows(I).Cells(7).Value.ToString.Trim = "03" Then

                    DataGridView1.Rows(I).DefaultCellStyle.BackColor = Color.DarkTurquoise
                    DataGridView1.Rows(I).DefaultCellStyle.ForeColor = Color.White

                    DataGridView1.Rows(I).Cells(8).Value = "TERMINADO"
                Else
                    DataGridView1.Rows(I).Cells(8).Value = "PENDIENTE"
                End If

            Next
            DataGridView1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(7).Visible = False
            DataGridView1.Columns(9).Visible = False
        Else
            MsgBox("LA OP INGRESADA NO EXISTE")
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If e.ColumnIndex = 0 Then

            ' MsgBox(("Row : " + e.RowIndex.ToString & "  Col : ") + e.ColumnIndex.ToString)'
            Dim num1, num2 As Integer
            num1 = e.RowIndex.ToString
            num2 = e.ColumnIndex.ToString
            Detalle_Confeccion.TextBox1.Text = DataGridView1.Rows(num1).Cells(9).Value.ToString.Trim
            Detalle_Confeccion.TextBox2.Text = DataGridView1.Rows(num1).Cells(3).Value.ToString.Trim
            Detalle_Confeccion.TextBox3.Text = DataGridView1.Rows(num1).Cells(5).Value.ToString.Trim
            Detalle_Confeccion.Show()
        End If
    End Sub

    Private Sub Confeccion_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class