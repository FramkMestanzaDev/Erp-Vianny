Public Class Cancelacion_documento
    Dim dt, dt2, dt3 As New DataTable
    Private Sub buscar()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(dt.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "FACTURA" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv

            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Cancelacion_documento_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim fr As New freportedoc

        dt = fr.reporte_tesoreria()
        If dt.Rows.Count <> 0 Then
            DataGridView1.DataSource = dt
            DataGridView1.Columns(1).Width = 120
            DataGridView1.Columns(2).Width = 120
            DataGridView1.Columns(3).Width = 100
            DataGridView1.Columns(4).Width = 120
            DataGridView1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(4).HeaderText = "ESTADO"
            Dim r As New Integer
            r = DataGridView1.Rows.Count
            For i = 0 To r - 1
                If DataGridView1.Rows(i).Cells(4).Value.ToString.Trim = 1 Then
                    DataGridView1.Rows(i).Cells(4).Value = "PENDIENTE"
                    DataGridView1.Columns(4).DefaultCellStyle.ForeColor = Color.Coral
                End If

            Next

        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim i As Integer
        Dim func As New vreportedoc
        Dim fg As New freportedoc
        i = DataGridView1.Rows.Count


        For i1 = 0 To i - 1
            If DataGridView1.Rows(i1).Cells(0).Value = True Then
                func.gestado = DataGridView1.Rows(i1).Cells(1).Value
                fg.actualizar_ESTADO(func)
            End If

        Next
        MsgBox("SE CANCELO LA FACTURA")
        dt = fg.reporte_tesoreria()

        DataGridView1.DataSource = dt
            DataGridView1.Columns(1).Width = 120
            DataGridView1.Columns(2).Width = 120
            DataGridView1.Columns(3).Width = 100
            DataGridView1.Columns(4).Width = 120
            DataGridView1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(4).HeaderText = "ESTADO"
            Dim r As New Integer
            r = DataGridView1.Rows.Count
            For i = 0 To r - 1
                If DataGridView1.Rows(i).Cells(4).Value.ToString.Trim = 1 Then
                    DataGridView1.Rows(i).Cells(4).Value = "PENDIENTE"
                    DataGridView1.Columns(4).DefaultCellStyle.ForeColor = Color.Coral
                End If

            Next


    End Sub
End Class