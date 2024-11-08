Public Class Buscar_Factura
    Dim DT As New DataTable
    Private Sub Buscar_Factura_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim GH As New freportedoc

        DT = GH.reportedoc5()
        If DT.Rows.Count <> 0 Then
            DataGridView1.DataSource = DT
            DataGridView1.Columns(1).Width = 200
            DataGridView1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        End If
    End Sub
    Private Sub buscar()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(DT.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "FACTURA" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(1).Width = 200
                DataGridView1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim i, a, conta As Integer
        i = DataGridView1.Rows.Count
        a = 0
        conta = 0



        For a = 0 To i - 1
            If Me.DataGridView1.Rows(a).Cells(0).Value = True Then
                conta = conta + 1
            End If
        Next
        If conta > 1 Then
            MsgBox("SOLO DEBE SELECCIONAR UNA OPCION")
        Else

            For a = 0 To i - 1
                If Me.DataGridView1.Rows(a).Cells(0).Value = True Then

                    Reporte_Facturas.TextBox1.Text = DataGridView1.Rows(a).Cells(1).Value

                End If
            Next

            Close()
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar()
    End Sub
End Class