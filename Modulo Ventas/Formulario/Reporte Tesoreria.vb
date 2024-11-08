Public Class Reporte_Tesoreria
    Dim dt As New DataTable
    Private Sub buscar()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(dt.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "FECHA" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv

            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub buscar1()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(dt.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "lacrado" & " like '%" & TextBox2.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv

            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub buscar2()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(dt.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "VENCIMIENTO" & " like '%" & TextBox3.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv

            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub Reporte_Tesoreria_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim kj As New freportedoc
        dt = kj.reporte_tesoreria2()
        DataGridView1.DataSource = dt

        DataGridView1.Columns(1).Width = 120
        DataGridView1.Columns(2).Width = 120
        DataGridView1.Columns(3).Width = 100

        DataGridView1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        DataGridView1.Columns(3).HeaderText = "ESTADO"
        Dim r As New Integer
        r = DataGridView1.Rows.Count
        For i = 0 To r - 1
            DataGridView1.Rows(i).Cells(4).Value = DateTime.Now.ToString("dd/MM/yyyy")
            If DataGridView1.Rows(i).Cells(3).Value.ToString.Trim = 1 Then
                DataGridView1.Rows(i).Cells(3).Value = "PENDIENTE"
                DataGridView1.Columns(3).DefaultCellStyle.ForeColor = Color.Indigo
            Else
                DataGridView1.Rows(i).Cells(3).Value = "CANCELADO"

            End If
            If DataGridView1.Rows(i).Cells(4).Value > DataGridView1.Rows(i).Cells(2).Value Then
                DataGridView1.Rows(i).Cells(5).Value = "VENCIDO"
                DataGridView1.Columns(5).DefaultCellStyle.BackColor = Color.DarkRed
                DataGridView1.Columns(5).DefaultCellStyle.ForeColor = Color.White
            Else
                DataGridView1.Rows(i).Cells(5).Value = "POR VENCER"
            End If
            DataGridView1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Next
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar()
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        buscar1()
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        buscar2()
    End Sub
End Class